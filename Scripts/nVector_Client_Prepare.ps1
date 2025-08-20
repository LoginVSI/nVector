# version 1.0.0
# -----------------------------
# Variables and Configuration
# -----------------------------

# How often nVector-agent polls for latency data (ms)
$NvectorAgentCheckIntervalMs = 5000  

# Full paths for CSV and logs
$CsvFilePath         = "C:\temp\nvidia\latency_metrics.csv"
$NvectorScreenshotDir = "C:\temp\nvidia\SSIM_screenshots"
$NvectorLogFile        = "C:\temp\nvidia\agent.log"

# --------------------------------------
# Start a transcript to capture ALL streams
# --------------------------------------
$Timestamp            = (Get-Date).ToString('yyyyMMddTHHmmss')
$TranscriptFile       = "C:\temp\nvidia\${Timestamp}nVector_Agent_Client.log"
$ErrorActionPreference = 'Continue'
$VerbosePreference     = 'Continue'
$DebugPreference       = 'Continue'
Start-Transcript -Path $TranscriptFile -Append

# Arguments for nVector-agent (client mode)
$Arguments = @(
    "-r","client",
    "-m", $CsvFilePath,
    "-p", "$NvectorAgentCheckIntervalMs",
    "-s", $NvectorScreenshotDir,
    "-l", $NvectorLogFile
)

# nVector-agent executable path
$NvectorAgentExePath = "" # Place full path to nvector-agent.exe here

# Polling intervals and thresholds
$PollingInterval     = 10      # seconds between CSV scans
$MaxLatencyThreshold = 1500    # ms, exclude outliers

# Launcher process details
$LauncherProcessName = "LoginEnterprise.Launcher.UI"
$LauncherExePath     = "C:\Program Files\Login VSI\Login Enterprise Launcher\LoginEnterprise.Launcher.UI.exe"

# API configuration
$ConfigurationAccessToken = "abcd1234abcd1234abcd1234abcd1234abcd1234abc"
$BaseUrl                = "https://myDomain.LoginEnterprise.com/"
$ApiEndpoint            = "publicApi/v7-preview/platform-metrics"
$EnvironmentId          = "abcd1234-abcd1234-abcd1234-abcd1"
$MetricId               = "nVectorMetricId"
$DisplayName            = "Endpoint Latency"
$Unit                   = "Latency (ms)"
$Instance               = $env:COMPUTERNAME
$Group                  = "nVector"
$ComponentType          = "vm"

# --------------------------------------
# Time-offset: pull server clock + compute clock drift with RTT compensation
# --------------------------------------
function Get-ServerOffset {
    param($BaseUrl, $ApiEndpoint, $Token)
    $uri = $BaseUrl.TrimEnd('/') + '/' + $ApiEndpoint.TrimStart('/')
    $hdr = @{ Authorization = "Bearer $Token" }

    try {
        $localBefore = [DateTimeOffset]::Now
        $resp        = Invoke-WebRequest -Uri $uri -Headers $hdr -UseBasicParsing -TimeoutSec 5
        $localAfter  = [DateTimeOffset]::Now

        $serverDto = [DateTimeOffset]::Parse($resp.Headers['Date'])

        # Compute half the round-trip
        $roundtrip = $localAfter - $localBefore
        $halfRt    = [TimeSpan]::FromTicks([Math]::Floor($roundtrip.Ticks / 2))

        # Estimate local time at midpoint
        $estimatedLocal = $localBefore + $halfRt

        $serverUtc   = $serverDto.UtcDateTime
        $localUtcEst = $estimatedLocal.UtcDateTime

        Write-Host "Server time (UTC):      $($serverUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
        Write-Host "Local before request:   $($localBefore.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
        Write-Host "Local after  request:   $($localAfter.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
        Write-Host "Estimated local (mid):  $($localUtcEst.ToString('yyyy-MM-ddTHH:mm:ssZ'))"

        $drift = $localUtcEst - $serverUtc
        Write-Host "Determined drift:       $($drift.ToString())"

        return $drift
    } catch {
        Write-Error "Failed to fetch server time from ${uri}: $($_.Exception.Message)"
        Stop-Transcript
        exit 1
    }
}

# Compute clock drift once at startup
$TimeOffsetSpan = Get-ServerOffset `
    -BaseUrl $BaseUrl `
    -ApiEndpoint "v8-preview/system/version" `
    -Token $ConfigurationAccessToken

# --------------------------------------
# Adjust each CSV timestamp and log details
# --------------------------------------
function Adjust-TimeOffset {
    param([string]$RawTimestamp)
    try {
        $dt       = [DateTime]::Parse($RawTimestamp)
        $adjusted = $dt.Add(-$TimeOffsetSpan)
        $finalStr = $adjusted.ToString("yyyy-MM-ddTHH:mm:ss") + "Z"

        Write-Host "Adjust-Timestamp → Raw: $RawTimestamp | Drift: $TimeOffsetSpan | Adjusted: $finalStr"
        return $finalStr
    } catch {
        Write-Host "Invalid timestamp format: $RawTimestamp"
        return $null
    }
}

function Terminate-nvector-agent {
    $p = Get-Process -Name "nvector-agent" -ErrorAction SilentlyContinue
    if ($p) { Write-Host "Killing previous nvector-agent"; $p | Stop-Process -Force }
    else  { Write-Host "nvector-agent not running" }
}

function Start-nvector-agent {
    if (-not (Test-Path $NvectorAgentExePath)) {
        Write-Error "nvector-agent.exe not found at $NvectorAgentExePath"; exit 1
    }
    Start-Process -FilePath $NvectorAgentExePath -ArgumentList $Arguments -NoNewWindow -ErrorAction Stop
    Write-Host "nvector-agent started"
}

function Upload-DataToApi {
    param([array]$Metrics)
    $json = $Metrics | ConvertTo-Json -Depth 10 -Compress
    if (-not $json.TrimStart().StartsWith("[")) { $json = "[$json]" }
    Write-Host "POST payload: $json"
    $hdr = @{
        Authorization = "Bearer $ConfigurationAccessToken"
        "Content-Type" = "application/json"
    }
    try {
        Invoke-RestMethod -Uri ($BaseUrl.TrimEnd('/') + '/' + $ApiEndpoint.TrimStart('/')) `
                          -Method Post -Headers $hdr -Body $json | Out-Null
        Write-Host "Upload succeeded"
    } catch {
        Write-Host "Upload error: $_"
    }
}

# -----------------------------
# Main Script Execution
# -----------------------------
Write-Host "Starting nVector metrics uploader"
Terminate-nvector-agent
Start-nvector-agent

# Ensure Launcher
if (-not (Get-Process -Name $LauncherProcessName -ErrorAction SilentlyContinue)) {
    Start-Process -FilePath $LauncherExePath -NoNewWindow -ErrorAction SilentlyContinue
    Write-Host "Launcher started"
} else {
    Write-Host "Launcher already running"
}

# Await CSV and header
$exists = $false
for ($i = 0; $i -lt 5; $i++) {
    if (Test-Path $CsvFilePath) { $exists = $true; break }
    Start-Sleep -Seconds 1
}
if (-not $exists) {
    Write-Host "CSV not found after wait"
} else {
    $headerLine = Get-Content $CsvFilePath -First 1
    if (-not $headerLine) {
        Write-Host "CSV exists but is empty"
    } elseif ($headerLine.Trim() -ne "timestamp,latency_ms") {
        Write-Host "Expected header 'timestamp,latency_ms' but found: '$headerLine'"
    } else {
        Write-Host "CSV ready"
    }
}

# Initialize line counter
if (Test-Path $CsvFilePath) {
    $LastLine = (Get-Content $CsvFilePath).Count - 1
} else {
    $LastLine = 0
}

# Watch loop
while ($true) {
    if (Test-Path $CsvFilePath) {
        Start-Sleep -Milliseconds 500
        $all = Get-Content $CsvFilePath
        if ($all.Count -le 1) { Start-Sleep $PollingInterval; continue }

        $current = $all.Count - 1
        if ($current -gt $LastLine) {
            $newLines = $all[($LastLine + 1)..$current]
            $metrics  = @()

            foreach ($line in $newLines) {
                $parts = $line -split ','
                if ($parts.Count -ne 2) {
                    Write-Host "Bad CSV line: $line"
                    continue
                }
                $tsRaw = $parts[0].Trim()
                $lat   = $parts[1].Trim()

                [double]$val = 0.0
                if ([double]::TryParse($lat, [ref]$val) -and $val -lt $MaxLatencyThreshold) {
                    $ts = Adjust-TimeOffset -RawTimestamp $tsRaw
                    if (-not $ts) { continue }
                    $metrics += [PSCustomObject]@{
                        metricId       = $MetricId
                        environmentKey = $EnvironmentId
                        timestamp      = $ts
                        displayName    = $DisplayName
                        unit           = $Unit
                        instance       = $Instance
                        value          = $val
                        group          = $Group
                        componentType  = $ComponentType
                    }
                } else {
                    Write-Host "Excluded or invalid latency: '$lat'"
                }
            }

            if ($metrics.Count) {
                Upload-DataToApi -Metrics $metrics
            }
            $LastLine = $current
        }
    }
    Start-Sleep -Seconds $PollingInterval
}

# -----------------------------
# Clean up transcript
# -----------------------------
Stop-Transcript
