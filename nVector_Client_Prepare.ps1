# -----------------------------
# Variables and Configuration
# -----------------------------

# How often nVector-agent polls for latency data (ms)
$NvectorAgentCheckIntervalMs = 5000  

# Full paths for CSV and logs
$CsvFilePath        = "C:\temp\nvidia\latency_metrics.csv"
$NvectorScreenshotDir = "C:\temp\nvidia\SSIM_screenshots"
$NvectorLogFile       = "C:\temp\nvidia\agent.log"
$ScriptLogFile        = "C:\temp\nvidia\nVector Prepare.txt"

# Arguments for nVector-agent (client mode)
$Arguments = @(
    "-r","client",
    "-m", $CsvFilePath,
    "-p", "$NvectorAgentCheckIntervalMs",
    "-s", $NvectorScreenshotDir,
    "-l", $NvectorLogFile
)

# -----------------------------
# Additional Variables
# -----------------------------

# nVector-agent executable path
$NvectorAgentExePath = "" # Place the full path to the nvector-agent.exe here

# Polling intervals and thresholds
$PollingInterval       = 10     # How often this script checks for new CSV lines (in seconds)
$MaxLatencyThreshold   = 10000  # Exclude latencies above 10s as spurious outliers

# CSV existence check parameters
$CsvCheckTimeoutSeconds = 5  # How many seconds we wait for the CSV file to appear
$CsvCheckIntervalSeconds = 1 # How often we re-check (in seconds) within that time

# Time offset configuration (UTC offset)
$TimeOffset = "0:00"  # Offset from UTC in hours:minutes, e.g., "-7:00" (PST) or "+2:00" (CEST)

# Launcher process details
$LauncherProcessName = "LoginEnterprise.Launcher.UI"
$LauncherExePath     = "C:\Program Files\Login VSI\Login Enterprise Launcher\LoginEnterprise.Launcher.UI.exe"

# API configuration
$ConfigurationAccessToken = "abcd1234abcd1234abcd1234abcd1234abcd1234abc" # The Login Enterprise configuration access token goes here 
$BaseUrl      = "https://myDomain.LoginEnterprise.com/" # The Login Enterprise base URL goes here
$ApiEndpoint  = "publicApi/v7-preview/platform-metrics"
$EnvironmentId= "abcd1234-abcd1234-abcd1234-abcd1"
$MetricId     = "nVectorMetricId"
$DisplayName  = "nVectorDisplayName"
$Unit         = "Latency"
$Instance     = "nVectorInstanceName"
$Group        = "nVectorGroup"
$ComponentType= $env:COMPUTERNAME

# -----------------------------
# Helper Functions
# -----------------------------

function Log {
    param([string]$Message)
    $Timestamp   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $LogMessage  = "$Timestamp $Message"
    Write-Host $LogMessage
    Add-Content -Path $ScriptLogFile -Value $LogMessage
}

function Adjust-TimeOffset {
    param([string]$RawTimestamp)
    try {
        $Datetime = [datetime]::Parse($RawTimestamp)
        $Offset   = [timespan]::Parse($TimeOffset)
        $Adjusted = $Datetime.Add($Offset)
        return $Adjusted.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    } catch {
        Log "Invalid timestamp format in line: $RawTimestamp"
        return $null
    }
}

function Terminate-nvector-agent {
    $ProcessName = "nvector-agent"
    $Running     = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($Running) {
        Log "Terminating all running '$ProcessName' processes."
        $Running | Stop-Process -Force -ErrorAction SilentlyContinue
    } else {
        Log "'$ProcessName' is not running."
    }
}

function Start-nvector-agent {
    if (-not (Test-Path $NvectorAgentExePath)) {
        Log "Executable not found at path: $NvectorAgentExePath"
        throw "Executable not found: $NvectorAgentExePath"
    }

    try {
        Start-Process -FilePath $NvectorAgentExePath -ArgumentList $Arguments -NoNewWindow -ErrorAction Stop
        Log "'nvector-agent' started successfully."
    } catch {
        Log "Failed to start 'nvector-agent'. Error: $_"
        throw
    }
}

function Upload-DataToApi {
    param([array]$MetricsArray)

    foreach ($m in $MetricsArray) {
        if (-not $m.PSObject.Properties['componentType']) {
            $m | Add-Member -MemberType NoteProperty -Name "componentType" -Value $ComponentType -Force
            Log "Added missing 'componentType' to metric: $($m | ConvertTo-Json -Compress)"
        }
    }

    $PayloadJson = $MetricsArray | ConvertTo-Json -Depth 10 -Compress
    if (-not ($PayloadJson.TrimStart().StartsWith("["))) {
        $PayloadJson = "[$PayloadJson]"
    }

    Log "Payload JSON: $PayloadJson"

    $Headers = @{ 
        "Authorization" = "Bearer $ConfigurationAccessToken"
        "Content-Type"  = "application/json" 
    }

    try {
        $FullUrl = "$($BaseUrl.TrimEnd('/'))/$($ApiEndpoint.TrimStart('/'))"
        Log "Sending POST request to $FullUrl ..."
        $Response = Invoke-RestMethod -Uri $FullUrl -Method Post -Headers $Headers -Body $PayloadJson
        Log "Response: $($Response | ConvertTo-Json -Depth 10)"
    } catch {
        Log "Error during API request: $_"
    }
}

# -----------------------------
# Main Script Execution
# -----------------------------

Log "Starting nVector metrics uploader script."
Terminate-nvector-agent
Start-nvector-agent

# Ensure the Launcher is running
if (-not (Get-Process -Name $LauncherProcessName -ErrorAction SilentlyContinue)) {
    try {
        Start-Process -FilePath $LauncherExePath -NoNewWindow -ErrorAction Stop
        Log "Launcher process '$LauncherProcessName' started successfully."
    } catch {
        Log "Failed to start launcher process '$LauncherProcessName'."
    }
} else {
    Log "Launcher process '$LauncherProcessName' is already running."
}

Log "Checking for CSV file existence and header..."

# Wait up to $CsvCheckTimeoutSeconds for CSV file to appear (checking once per second)
$csvExists = $false
for ($i = 1; $i -le $CsvCheckTimeoutSeconds; $i++) {
    if (Test-Path $CsvFilePath) {
        $csvExists = $true
        break
    }
    Start-Sleep -Seconds $CsvCheckIntervalSeconds
}

if (-not $csvExists) {
    Log "CSV file '$CsvFilePath' does not exist after waiting $CsvCheckTimeoutSeconds seconds."
} else {
    # Check the first line is the expected header
    $firstLine = (Get-Content -Path $CsvFilePath | Select-Object -First 1).Trim()
    if ($firstLine -ne "timestamp,latency_ms") {
        Log "Expected header 'timestamp,latency_ms' but found: '$firstLine'"
    } else {
        Log "CSV file found and header is correct."
    }
}

Log "Monitoring CSV file '$CsvFilePath' for new lines and uploading data to API..."

# Initialize $LastLineCount
if (Test-Path -Path $CsvFilePath) {
    $Lines = Get-Content -Path $CsvFilePath
    $LastLineCount = $Lines.Count - 1
} else {
    $LastLineCount = 0
    Log "CSV file does not exist yet."
}

# Continuous loop to watch for new data
while ($true) {
    if (Test-Path $CsvFilePath) {
        try {
            # Small sleep to avoid partial line reads
            Start-Sleep -Milliseconds 500
            $Lines = Get-Content -Path $CsvFilePath -ErrorAction Stop

            # If we only have the header (or zero lines), just wait and continue
            if ($Lines.Count -le 1) {
                Start-Sleep -Seconds $PollingInterval
                continue
            }

            $CurrentDataCount = $Lines.Count - 1
            if ($CurrentDataCount -gt $LastLineCount) {
                $NewDataCount = $CurrentDataCount - $LastLineCount
                $NewLines = $Lines[($LastLineCount + 1)..($LastLineCount + $NewDataCount)]
                $MetricsArray = @()

                foreach ($NewLine in $NewLines) {
                    $Parts = $NewLine -split ','
                    if ($Parts.Count -lt 2) {
                        Log "Invalid line format: $NewLine"
                        continue
                    }

                    $timeRaw = $Parts[0].Trim()
                    $latStr  = $Parts[1].Trim()
                    $UtcTime = Adjust-TimeOffset -RawTimestamp $timeRaw
                    if (-not $UtcTime) { continue }

                    [float]$latVal = 0.0
                    if ([float]::TryParse($latStr, [ref]$latVal)) {
                        if ($latVal -lt $MaxLatencyThreshold) {
                            $Metric = [PSCustomObject]@{
                                metricId       = $MetricId
                                environmentKey = $EnvironmentId
                                timestamp      = $UtcTime
                                displayName    = $DisplayName
                                unit           = $Unit
                                instance       = $Instance
                                value          = $latVal
                                group          = $Group
                                componentType  = $ComponentType
                            }
                            $MetricsArray += $Metric
                        } else {
                            Log "Excluded high latency: $latVal ms."
                        }
                    } else {
                        Log "Invalid latency value in line: $NewLine"
                    }
                }

                if ($MetricsArray.Count -gt 0) {
                    Upload-DataToApi -MetricsArray $MetricsArray
                }

                $LastLineCount = $CurrentDataCount
            }
        } catch {
            Log "Error processing CSV file: $_"
        }
    }

    Start-Sleep -Seconds $PollingInterval
}