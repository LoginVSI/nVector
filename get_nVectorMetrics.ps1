<#
.SYNOPSIS
Retrieves metrics from the Login Enterprise v7-preview/platform-metrics endpoint and saves the JSON and CSV outputs to %TEMP%, bypassing SSL certificate validation as needed.

.DESCRIPTION
This script:
1. **Bypasses SSL certificate validation**:
   - On PowerShell 7.x: uses `-SkipCertificateCheck` with `Invoke-RestMethod`.
   - On PowerShell 5.x: sets `[System.Net.ServicePointManager]::ServerCertificateValidationCallback` to always return `$true`.
2. **Accepts mandatory parameters** for `StartTime`, `EndTime`, and `EnvironmentId`, all in ISO 8601 Zulu format (e.g., `2025-02-07T15:21:53.346Z`).
3. **Accepts optional parameters**:
   - `ApiAccessToken` (overrides the default token in the script)
   - `BaseUrl`        (overrides the default base URL in the script)
   - `OutputCsvFilePath` (defaults to `%TEMP%\get_nVectorMetrics.csv` if omitted)
   - `OutputJsonFilePath` (defaults to `%TEMP%\get_nVectorMetrics.json` if omitted)
4. **Saves a log** of all non-error and error messages to `%TEMP%\get_nVectorMetrics_Log.txt`, **excluding** the raw CSV/JSON content.
5. **Behavior for Console Output**:
   - If `-OutputJsonFilePath` is not specified, the raw JSON is shown in the console; otherwise, it’s **not** displayed, only saved to a file.
   - If `-OutputCsvFilePath` is not specified, the CSV data is shown in the console; otherwise, it’s **not** displayed, only saved to a file.
6. **Preserves** timestamps in the JSON/CSV as ISO 8601 Z.

.NOTES
- **Bypassing certificate validation is insecure** for production use.
- For large datasets, be mindful of memory/time requirements when converting JSON to CSV.

.PARAMETER StartTime
(Mandatory) Start of the date/time range (ISO 8601 Z). Example: `2025-02-07T15:21:53.346Z`

.PARAMETER EndTime
(Mandatory) End of the date/time range (ISO 8601 Z). Example: `2025-02-07T17:00:00.000Z`

.PARAMETER EnvironmentId
(Mandatory) The Login Enterprise environment ID to query.

.PARAMETER ApiAccessToken
(Optional) Overrides the default token if provided.

.PARAMETER BaseUrl
(Optional) Overrides the default base URL if provided.

.PARAMETER OutputCsvFilePath
(Optional) If omitted, defaults to `%TEMP%\get_nVectorMetrics.csv`.  
If specified, the CSV content is **not** shown in the console.

.PARAMETER OutputJsonFilePath
(Optional) If omitted, defaults to `%TEMP%\get_nVectorMetrics.json`.  
If specified, the JSON content is **not** shown in the console.

.PARAMETER Help
Displays usage information.

.EXAMPLE
# 1) Using only mandatory parameters:
PS C:\> .\Get-nVectorMetrics.ps1 -StartTime "2025-02-07T00:00:00.000Z" -EndTime "2025-02-07T23:59:59.999Z" -EnvironmentId "abcdef1234"

Outputs JSON and CSV content in the console, and also writes them to `%TEMP%\get_nVectorMetrics.json/.csv` by default. Logs are at `%TEMP%\get_nVectorMetrics_Log.txt`.

.EXAMPLE
# 2) Providing custom paths for JSON and CSV:
PS C:\> .\Get-nVectorMetrics.ps1 -StartTime "2025-02-07T00:00:00.000Z" -EndTime "2025-02-07T23:59:59.999Z" -EnvironmentId "abcdef1234" -ApiAccessToken "MY_TOKEN" -BaseUrl "https://mydomain.loginenterprise.com" -OutputCsvFilePath "C:\temp\nVector.csv" -OutputJsonFilePath "C:\temp\nVector.json"

Saves JSON to `C:\temp\nVector.json`, CSV to `C:\temp\nVector.csv`, and does **not** display either in the console. Logs are at `%TEMP%\get_nVectorMetrics_Log.txt`.
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$StartTime,

    [Parameter(Mandatory=$true)]
    [string]$EndTime,

    [Parameter(Mandatory=$true)]
    [string]$EnvironmentId,

    [string]$ApiAccessToken,
    [string]$BaseUrl,

    [string]$OutputCsvFilePath,
    [string]$OutputJsonFilePath,

    [switch]$Help
)

# ---------------------------------------------------------------------
# Set up default file paths in %TEMP% if not provided
# ---------------------------------------------------------------------
$TempPath = $env:TEMP
if (-not $OutputCsvFilePath)   { $OutputCsvFilePath   = Join-Path $TempPath "get_nVectorMetrics.csv" }
if (-not $OutputJsonFilePath)  { $OutputJsonFilePath  = Join-Path $TempPath "get_nVectorMetrics.json" }
$ScriptLogFile = Join-Path $TempPath "get_nVectorMetrics_Log.txt"

# ---------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------
function Write-Log {
    param([string]$Message, [switch]$IsError)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formatted = "$timestamp - $Message"
    Add-Content -Path $ScriptLogFile -Value $formatted

    # Only write to console if it's an error
    if ($IsError) {
        Write-Error $Message
    }
}

Write-Log "==== Script invoked. ===="

# ---------------------------------------------------------------------
# HELP / USAGE
# ---------------------------------------------------------------------
if ($Help) {
    Write-Host "Usage:"
    Write-Host "  .\Get-nVectorMetrics.ps1 -StartTime <ISO8601Z> -EndTime <ISO8601Z> -EnvironmentId <ID>"
    Write-Host "                         [-ApiAccessToken <Token>] [-BaseUrl <URL>] [-OutputCsvFilePath <Path>] [-OutputJsonFilePath <Path>]"
    Write-Host ""
    Write-Host "Mandatory Parameters:"
    Write-Host "  -StartTime          e.g. 2025-02-07T15:21:53.346Z"
    Write-Host "  -EndTime            e.g. 2025-02-07T17:00:00.000Z"
    Write-Host "  -EnvironmentId      The environment ID to query"
    Write-Host ""
    Write-Host "Optional Parameters:"
    Write-Host "  -ApiAccessToken     Overrides the default token"
    Write-Host "  -BaseUrl            Overrides the default base URL"
    Write-Host "  -OutputCsvFilePath  If not specified, defaults to $OutputCsvFilePath"
    Write-Host "  -OutputJsonFilePath If not specified, defaults to $OutputJsonFilePath"
    Write-Host "  -Help               Show this help"
    Write-Host ""
    Write-Host "Examples:"
    Write-Host '  PS> .\Get-nVectorMetrics.ps1 -StartTime "2025-02-07T00:00:00.000Z" -EndTime "2025-02-07T23:59:59.999Z" -EnvironmentId "abcdef1234"'
    Write-Host '  PS> .\Get-nVectorMetrics.ps1 -StartTime "2025-02-07T00:00:00.000Z" -EndTime "2025-02-07T23:59:59.999Z" -EnvironmentId "abcdef1234" -ApiAccessToken "MY_TOKEN" -BaseUrl "https://mydomain.loginenterprise.com" -OutputCsvFilePath "C:\temp\nVector.csv" -OutputJsonFilePath "C:\temp\nVector.json"'
    Write-Host ""
    Write-Host "Note: SSL certificate validation is bypassed in this script. Logs are at $ScriptLogFile"
    return
}

# ---------------------------------------------------------------------
# DEFAULTS
# ---------------------------------------------------------------------
$DefaultApiAccessToken = "YOUR-DEFAULT-TOKEN-GOES-HERE"
$DefaultBaseUrl        = "https://myDomain.LoginEnterprise.com/"

# ---------------------------------------------------------------------
# CERTIFICATE VALIDATION BYPASS
# ---------------------------------------------------------------------
if ($PSVersionTable.PSVersion.Major -ge 7) {
    Write-Log "PowerShell 7.x detected -> will use -SkipCertificateCheck."
} else {
    Write-Log "PowerShell <7 -> bypassing certificate validation via ServerCertificateValidationCallback."
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }
}

# ---------------------------------------------------------------------
# VALIDATE ISO 8601 Z FORMAT
# ---------------------------------------------------------------------
function Validate-IsIso8601Zulu {
    param([string]$DateTimeStr)
    return ($DateTimeStr -match "^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z$")
}

if (-not (Validate-IsIso8601Zulu $StartTime)) {
    Write-Log "Invalid StartTime format. Must be ISO8601Z." -IsError
    return
}
if (-not (Validate-IsIso8601Zulu $EndTime)) {
    Write-Log "Invalid EndTime format. Must be ISO8601Z." -IsError
    return
}

# ---------------------------------------------------------------------
# DETERMINE TOKEN & BASE URL
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
# BUILD GET URL
# ---------------------------------------------------------------------
$queryParams = "?from=$StartTime&to=$EndTime&environmentIds=$EnvironmentId"
$FullUrl = "$UsedBaseUrl/publicApi/v7-preview/platform-metrics$queryParams"
Write-Log "Constructed URL: $FullUrl"

$Headers = @{
    "Authorization" = "Bearer $UsedApiAccessToken"
    "Accept"        = "application/json"
}

# ---------------------------------------------------------------------
# INVOKE GET REQUEST
# ---------------------------------------------------------------------
Write-Log "Sending GET request..."
try {
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $JsonResponse = Invoke-RestMethod -Uri $FullUrl -Method GET -Headers $Headers -SkipCertificateCheck -ErrorAction Stop
    } else {
        $JsonResponse = Invoke-RestMethod -Uri $FullUrl -Method GET -Headers $Headers -ErrorAction Stop
    }
    Write-Log "GET request succeeded."
} catch {
    Write-Log "Error during GET request: $_" -IsError
    return
}

# ---------------------------------------------------------------------
# HANDLE JSON OUTPUT
# ---------------------------------------------------------------------
$JsonToSave = $JsonResponse | ConvertTo-Json -Depth 10
try {
    $JsonToSave | Out-File -FilePath $OutputJsonFilePath -Encoding UTF8
    Write-Log "JSON saved to: $OutputJsonFilePath"
} catch {
    Write-Log "Failed to write JSON output: $_" -IsError
}

# Show JSON in console ONLY if user did NOT specify -OutputJsonFilePath
# (Because user requested to hide JSON from console if they explicitly provide a file path)
if ($PSBoundParameters.ContainsKey('OutputJsonFilePath')) {
    Write-Log "Not displaying JSON in console because user specified -OutputJsonFilePath."
} else {
    Write-Host "`n--- JSON Response (Start) ---"
    Write-Host $JsonToSave
    Write-Host "--- JSON Response (End) ---`n"
}

# ---------------------------------------------------------------------
# CONVERT TO CSV
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
            # Preserve raw string format for the timestamp
            $finalTimestamp = $dp.timestamp
            $row = [PSCustomObject]@{
                timestamp      = $finalTimestamp
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
    $CsvData = $AllDataRows | ConvertTo-Csv -NoTypeInformation
    # Save CSV to file always
    try {
        $CsvData | Out-File -FilePath $OutputCsvFilePath -Encoding UTF8
        Write-Log "CSV saved to: $OutputCsvFilePath"
    } catch {
        Write-Log "Failed to write CSV output: $_" -IsError
    }

    # Show CSV in console ONLY if user did NOT specify -OutputCsvFilePath
    if ($PSBoundParameters.ContainsKey('OutputCsvFilePath')) {
        Write-Log "Not displaying CSV in console because user specified -OutputCsvFilePath."
    } else {
        Write-Host "`n--- CSV Output (Start) ---"
        Write-Host $CsvData
        Write-Host "--- CSV Output (End) ---`n"
    }
}

Write-Log "Script completed successfully."
