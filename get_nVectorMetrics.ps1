<#
.SYNOPSIS
Retrieves metrics from the Login Enterprise v7-preview/platform-metrics endpoint and saves the JSON and CSV outputs to files in C:\temp, bypassing SSL/TLS certificate validation. All events are logged to both the console and a log file (default C:\temp\get_nVectorMetrics_Log.txt).

.DESCRIPTION
This script:
1. Forces all available SSL/TLS protocols (Ssl3, Tls, Tls11, Tls12) and disables certificate validation.
2. Accepts mandatory parameters for StartTime, EndTime, and EnvironmentId (all in ISO 8601 Zulu format, e.g. 2025-02-07T15:21:53.346Z).
3. Accepts optional parameters for ApiAccessToken, BaseUrl, OutputCsvFilePath, OutputJsonFilePath, and LogFilePath.
   - If –OutputCsvFilePath is omitted, it defaults to C:\temp\get_nVectorMetrics.csv.
   - If –OutputJsonFilePath is omitted, it defaults to C:\temp\get_nVectorMetrics.json.
   - If –LogFilePath is omitted, it defaults to C:\temp\get_nVectorMetrics_Log.txt.
4. Logs all events to the specified log file and prints summary messages to the console—including the detected PowerShell version.
5. In PowerShell 7.x, uses Invoke-RestMethod with –SkipCertificateCheck; in PowerShell 5.x, uses HttpWebRequest.
6. Saves the retrieved JSON and converted CSV to the defined files without outputting the raw content to the console if an output file is specified.

.NOTES
*** INSECURE: Certificate validation is completely bypassed. ***
Use only in trusted environments.

.PARAMETER StartTime
(Mandatory) ISO 8601 Z start time, e.g. 2025-02-07T15:21:53.346Z.

.PARAMETER EndTime
(Mandatory) ISO 8601 Z end time, e.g. 2025-02-07T17:00:00.000Z.

.PARAMETER EnvironmentId
(Mandatory) Environment ID for filtering metrics.

.PARAMETER ApiAccessToken
(Optional) Overrides the default API token.

.PARAMETER BaseUrl
(Optional) Overrides the default base URL.

.PARAMETER OutputCsvFilePath
(Optional) CSV output file path; defaults to C:\temp\get_nVectorMetrics.csv.
If provided, raw CSV data is not displayed on the console.

.PARAMETER OutputJsonFilePath
(Optional) JSON output file path; defaults to C:\temp\get_nVectorMetrics.json.
If provided, raw JSON data is not displayed on the console.

.PARAMETER LogFilePath
(Optional) Log file path; defaults to C:\temp\get_nVectorMetrics_Log.txt.

.PARAMETER Help
Displays this help information.

.EXAMPLE
PS C:\> .\get_nVectorMetrics.ps1 -StartTime "2025-02-07T00:00:00.000Z" -EndTime "2025-02-07T23:59:59.999Z" -EnvironmentId "abcdef1234"

Uses default token and URL; JSON is saved to C:\temp\get_nVectorMetrics.json, CSV to C:\temp\get_nVectorMetrics.csv, and logs to C:\temp\get_nVectorMetrics_Log.txt.

.EXAMPLE
PS C:\> .\get_nVectorMetrics.ps1 -StartTime "2025-02-07T00:00:00.000Z" -EndTime "2025-02-07T23:59:59.999Z" -EnvironmentId "abcdef1234" -ApiAccessToken "MY_TOKEN" -BaseUrl "https://mydomain.com" -OutputCsvFilePath "C:\temp\nVector.csv" -OutputJsonFilePath "C:\temp\nVector.json" -LogFilePath "C:\temp\myLog.txt"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$StartTime,

    [Parameter(Mandatory = $true)]
    [string]$EndTime,

    [Parameter(Mandatory = $true)]
    [string]$EnvironmentId,

    [string]$ApiAccessToken,
    [string]$BaseUrl,

    [string]$OutputCsvFilePath,
    [string]$OutputJsonFilePath,
    [string]$LogFilePath,

    [switch]$Help
)

# ---------------------------------------------------------------------
# Set up default file paths (defaults in C:\temp)
# ---------------------------------------------------------------------
$DefaultCsvFilePath  = "C:\temp\get_nVectorMetrics.csv"
$DefaultJsonFilePath = "C:\temp\get_nVectorMetrics.json"
$DefaultLogFilePath  = "C:\temp\get_nVectorMetrics_Log.txt"

if (-not $OutputCsvFilePath)  { $OutputCsvFilePath  = $DefaultCsvFilePath }
if (-not $OutputJsonFilePath) { $OutputJsonFilePath = $DefaultJsonFilePath }
if (-not $LogFilePath)        { $LogFilePath        = $DefaultLogFilePath }
$ScriptLogFile = $LogFilePath

# ---------------------------------------------------------------------
# Logging function (only summary messages)
# ---------------------------------------------------------------------
function Write-Log {
    param([string]$Message, [switch]$IsError)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formatted = "$timestamp - $Message"
    Add-Content -Path $ScriptLogFile -Value $formatted
    Write-Host $formatted
    if ($IsError) { Write-Error $Message }
}

Write-Log "==== Script invoked. ===="

# ---------------------------------------------------------------------
# Show help if requested
# ---------------------------------------------------------------------
if ($Help) {
    Write-Host "Usage:"
    Write-Host "  .\get_nVectorMetrics.ps1 -StartTime <ISO8601Z> -EndTime <ISO8601Z> -EnvironmentId <ID>"
    Write-Host "                         [-ApiAccessToken <Token>] [-BaseUrl <URL>]"
    Write-Host "                         [-OutputCsvFilePath <Path>] [-OutputJsonFilePath <Path>]"
    Write-Host "                         [-LogFilePath <Path>] [-Help]"
    Write-Host ""
    Write-Host "Mandatory Parameters:"
    Write-Host "  -StartTime        e.g. 2025-02-07T15:21:53.346Z"
    Write-Host "  -EndTime          e.g. 2025-02-07T17:00:00.000Z"
    Write-Host "  -EnvironmentId    e.g. abcdef1234"
    Write-Host ""
    Write-Host "Optional Parameters:"
    Write-Host "  -ApiAccessToken     Overrides default token."
    Write-Host "  -BaseUrl            Overrides default base URL."
    Write-Host "  -OutputCsvFilePath  Defaults to $DefaultCsvFilePath"
    Write-Host "  -OutputJsonFilePath Defaults to $DefaultJsonFilePath"
    Write-Host "  -LogFilePath        Defaults to $DefaultLogFilePath"
    Write-Host "  -Help               Show this help."
    Write-Host ""
    Write-Host "Logs are saved to: $ScriptLogFile"
    Write-Host "** WARNING: Certificate validation is completely bypassed. **"
    return
}

# ---------------------------------------------------------------------
# Log detected PowerShell version
# ---------------------------------------------------------------------
$psVersion = $PSVersionTable.PSVersion
Write-Log "Detected PowerShell version: $($psVersion.Major).$($psVersion.Minor)"

# ---------------------------------------------------------------------
# Defaults for API token and BaseUrl
# ---------------------------------------------------------------------
$DefaultApiAccessToken = "YOUR-DEFAULT-TOKEN-GOES-HERE"
$DefaultBaseUrl        = "https://myDomain.LoginEnterprise.com"

# ---------------------------------------------------------------------
# Force SSL/TLS protocols and disable certificate validation
# ---------------------------------------------------------------------
try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
} catch {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
}
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }
Write-Log "SSL/TLS protocols forced; certificate validation bypassed."

# ---------------------------------------------------------------------
# Validate ISO 8601 Z format for StartTime and EndTime
# ---------------------------------------------------------------------
function Validate-IsIso8601Zulu {
    param([string]$DateTimeStr)
    return ($DateTimeStr -match "^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z$")
}
if (-not (Validate-IsIso8601Zulu $StartTime)) {
    Write-Log "Invalid StartTime format. Must be ISO8601Z: $StartTime" -IsError
    return
}
if (-not (Validate-IsIso8601Zulu $EndTime)) {
    Write-Log "Invalid EndTime format. Must be ISO8601Z: $EndTime" -IsError
    return
}

# ---------------------------------------------------------------------
# Determine final API token and BaseUrl
# ---------------------------------------------------------------------
if ($ApiAccessToken) {
    $UsedApiAccessToken = $ApiAccessToken
    Write-Log "Using user-provided API token."
} else {
    $UsedApiAccessToken = $DefaultApiAccessToken
    Write-Log "No -ApiAccessToken provided; using default token."
}
if ($BaseUrl) {
    $UsedBaseUrl = $BaseUrl.TrimEnd('/')
    Write-Log "Using user-provided BaseUrl: $UsedBaseUrl"
} else {
    $UsedBaseUrl = $DefaultBaseUrl.TrimEnd('/')
    Write-Log "No -BaseUrl provided; using default: $UsedBaseUrl"
}

Write-Log "Using CSV path: $OutputCsvFilePath"
Write-Log "Using JSON path: $OutputJsonFilePath"

# ---------------------------------------------------------------------
# Build GET URL
# ---------------------------------------------------------------------
$queryParams = "?from=$StartTime&to=$EndTime&environmentIds=$EnvironmentId"
$FullUrl = "$UsedBaseUrl/publicApi/v7-preview/platform-metrics$queryParams"
Write-Log "Constructed URL: $FullUrl"

$Headers = @{
    "Authorization" = "Bearer $UsedApiAccessToken"
    "Accept"        = "application/json"
}

# ---------------------------------------------------------------------
# Perform GET request based on PowerShell version
# ---------------------------------------------------------------------
if ($psVersion.Major -ge 7) {
    Write-Log "Using Invoke-RestMethod with -SkipCertificateCheck (PowerShell 7.x)."
    try {
        $jsonResult = Invoke-RestMethod -Uri $FullUrl -Method GET -Headers $Headers -SkipCertificateCheck -ErrorAction Stop
        Write-Log "GET request succeeded."
    } catch {
        Write-Log "Error during GET request: $_" -IsError
        return
    }
    # In PS 7, if the result is a string, convert it; otherwise assume it's already parsed.
    if ($jsonResult -is [string]) {
        try {
            $JsonResponse = $jsonResult | ConvertFrom-Json
        } catch {
            Write-Log "Failed to parse JSON response: $_" -IsError
            return
        }
    } else {
        $JsonResponse = $jsonResult
    }
    # For logging, get the JSON text version
    $jsonString = $JsonResponse | ConvertTo-Json -Depth 10
} else {
    Write-Log "Using HttpWebRequest (PowerShell 5.x)."
    try {
        $request = [System.Net.HttpWebRequest]::Create($FullUrl)
        $request.Method = "GET"
        $request.Headers.Add("Authorization", "Bearer $UsedApiAccessToken")
        $request.Accept = "application/json"
        $request.Timeout = 60000
        $response = $request.GetResponse()
        $stream = $response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream)
        $jsonString = $reader.ReadToEnd()
        $reader.Close()
        $response.Close()
        Write-Log "GET request succeeded."
    } catch {
        Write-Log "Error during GET request: $_" -IsError
        return
    }
    try {
        $JsonResponse = $jsonString | ConvertFrom-Json
        if (-not $JsonResponse) { throw "Empty or invalid JSON response." }
    } catch {
        Write-Log "Failed to parse JSON response: $_" -IsError
        return
    }
}

# ---------------------------------------------------------------------
# Save JSON to file (raw JSON text)
# ---------------------------------------------------------------------
try {
    $jsonString | Out-File -FilePath $OutputJsonFilePath -Encoding UTF8
    Write-Log "JSON saved to: $OutputJsonFilePath"
} catch {
    Write-Log "Failed to write JSON output: $_" -IsError
}

Write-Log "Raw JSON not displayed; saved to $OutputJsonFilePath."

# ---------------------------------------------------------------------
# Convert JSON to CSV (preserving raw timestamp strings)
# ---------------------------------------------------------------------
Write-Log "Converting JSON to CSV..."
$AllDataRows = @()

foreach ($metric in $JsonResponse) {
    $metricId       = $metric.metricId
    $environmentKey = $metric.environmentKey
    $displayName    = $metric.displayName
    $unit           = $metric.unit
    $instance       = $metric.instance
    $group          = $metric.group
    $componentType  = $metric.componentType

    if ($metric.dataPoints -and $metric.dataPoints.Count -gt 0) {
        foreach ($dp in $metric.dataPoints) {
            $row = [PSCustomObject]@{
                timestamp      = [string]$dp.timestamp
                value          = $dp.value
                metricId       = $metricId
                environmentKey = $environmentKey
                displayName    = $displayName
                unit           = $unit
                instance       = $instance
                group          = $group
                componentType  = $componentType
            }
            $AllDataRows += $row
        }
    }
}

if ($AllDataRows.Count -eq 0) {
    Write-Log "No metric data points found in the response."
} else {
    $CsvOutput = $AllDataRows | ConvertTo-Csv -NoTypeInformation
    try {
        $CsvOutput | Out-File -FilePath $OutputCsvFilePath -Encoding UTF8
        Write-Log "CSV saved to: $OutputCsvFilePath"
    } catch {
        Write-Log "Failed to write CSV output: $_" -IsError
    }
    Write-Log "Raw CSV not displayed; saved to $OutputCsvFilePath."
}

Write-Log "Script completed successfully."
