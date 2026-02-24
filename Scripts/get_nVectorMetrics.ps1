<#
.SYNOPSIS
    Login Enterprise nVector Platform Metrics Retrieval Tool

.DESCRIPTION
    Retrieves nVector Platform Metrics data (latency, SSIM, etc.) from the Login Enterprise API.
    Supports one or multiple Environment IDs in a single run.
    Exports results to timestamped CSV and JSON files for analysis.

.PARAMETER LEApiToken
    REQUIRED. Login Enterprise API token (Configuration access level).

.PARAMETER EnvironmentId
    Single Environment UUID. Use this OR -EnvironmentIds, not both.

.PARAMETER EnvironmentIds
    Array of Environment UUIDs. Use this OR -EnvironmentId, not both.
    Example: @("uuid-1", "uuid-2")

.PARAMETER StartTime
    Start of time range in ISO 8601 Zulu format. e.g. 2025-02-07T00:00:00.000Z
    If omitted, -LastHours is used instead.

.PARAMETER EndTime
    End of time range in ISO 8601 Zulu format. e.g. 2025-02-07T23:59:59.999Z
    If omitted, -LastHours is used instead.

.PARAMETER LastHours
    Convenience parameter. Retrieve metrics from the last N hours. Default: 1.
    Ignored if -StartTime and -EndTime are provided.

.PARAMETER BaseUrl
    Base URL of the Login Enterprise appliance. e.g. https://myDomain.LoginEnterprise.com

.PARAMETER ApiVersion
    API version segment. Default: v8-preview.
    Use v7-preview for older appliances.

.PARAMETER MetricGroups
    Optional array of metric group filters to narrow results.

.PARAMETER OutputDir
    Directory for output files. Defaults to script directory.
    Output filenames are auto-generated with timestamps.

.PARAMETER LogFilePath
    Path for script log file. Defaults to OutputDir\get_nVectorMetrics_Log_<timestamp>.txt

.PARAMETER ImportServerCert
    Import the appliance certificate into CurrentUser\Root before the request.
    Use for appliances with self-signed or private CA certificates.

.PARAMETER KeepCert
    Used with -ImportServerCert. Keeps imported certs after the run.
    If omitted, any newly imported certs are removed on exit.

.EXAMPLE
    # Last 1 hour, single environment
    .\get_nVectorMetrics.ps1 -LEApiToken "mytoken" -EnvironmentId "abcd-1234" -BaseUrl "https://my.le.com"

.EXAMPLE
    # Specific time range, multiple environments
    .\get_nVectorMetrics.ps1 -LEApiToken "mytoken" -EnvironmentIds @("uuid-1","uuid-2") -StartTime "2025-02-07T00:00:00.000Z" -EndTime "2025-02-07T23:59:59.999Z" -BaseUrl "https://my.le.com"

.EXAMPLE
    # Self-signed cert, last 2 hours
    .\get_nVectorMetrics.ps1 -LEApiToken "mytoken" -EnvironmentId "abcd-1234" -BaseUrl "https://appliance.local" -LastHours 2 -ImportServerCert

.EXAMPLE
    # Keep imported cert for future sessions
    .\get_nVectorMetrics.ps1 -LEApiToken "mytoken" -EnvironmentId "abcd-1234" -BaseUrl "https://appliance.local" -ImportServerCert -KeepCert

.NOTES
    Version:    2.0.0
    Author:     Login VSI
    Updated:    February 2026

    Security note: -ImportServerCert imports into CurrentUser\Root and -SkipCertificateCheck
    (PS7) bypass certificate validation. Use only on trusted/test networks.

    PS 5.x: Uses HttpWebRequest (cert validation not skipped). Use -ImportServerCert for
    appliances with self-signed/private CA certs.
    PS 7.x: Uses Invoke-RestMethod -SkipCertificateCheck (validation bypassed).
#>

param(
    [Parameter(Mandatory = $true)][string]$LEApiToken,
    [Parameter(Mandatory = $false)][string]$EnvironmentId,
    [Parameter(Mandatory = $false)][string[]]$EnvironmentIds,
    [Parameter(Mandatory = $false)][string]$StartTime,
    [Parameter(Mandatory = $false)][string]$EndTime,
    [Parameter(Mandatory = $false)][int]$LastHours = 1,
    [Parameter(Mandatory = $false)][string]$BaseUrl = "https://your-le-appliance.example.com",
    [Parameter(Mandatory = $false)][string]$ApiVersion = "v8-preview",
    [Parameter(Mandatory = $false)][string[]]$MetricGroups,
    [Parameter(Mandatory = $false)][string]$OutputDir,
    [Parameter(Mandatory = $false)][string]$LogFilePath,
    [Parameter(Mandatory = $false)][switch]$ImportServerCert,
    [Parameter(Mandatory = $false)][switch]$KeepCert
)

# =====================================================
# Version & Output Setup
# =====================================================
$ScriptVersion = "2.0.0"
$Timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")

if (-not $OutputDir) { $OutputDir = $PSScriptRoot }
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$CsvPath  = Join-Path $OutputDir "get_nVectorMetrics_$Timestamp.csv"
$JsonPath = Join-Path $OutputDir "get_nVectorMetrics_$Timestamp.json"
if (-not $LogFilePath) { $LogFilePath = Join-Path $OutputDir "get_nVectorMetrics_Log_$Timestamp.txt" }

$script:ImportedCertThumbs = @()

# =====================================================
# Logging
# =====================================================
function Write-Log {
    param([string]$Message, [switch]$IsError, [switch]$IsWarning)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formatted = "$ts - $Message"
    try {
        $logDir = Split-Path -Parent $LogFilePath
        if ($logDir -and -not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
        Add-Content -Path $LogFilePath -Value $formatted
    } catch {
        try { Write-Host ("WARNING: could not write to log file: {0}" -f $_.Exception.Message) -ForegroundColor Yellow } catch {}
    }
    if ($IsError)   { Write-Host $formatted -ForegroundColor Red }
    elseif ($IsWarning) { Write-Host $formatted -ForegroundColor Yellow }
    else            { Write-Host $formatted }
}

# =====================================================
# Banner
# =====================================================
Write-Host "`n========================================================================" -ForegroundColor Cyan
Write-Host "  Login Enterprise nVector Metrics Retrieval Tool v$ScriptVersion" -ForegroundColor Cyan
Write-Host "========================================================================`n" -ForegroundColor Cyan
Write-Log "==== Script started. Version $ScriptVersion ===="
Write-Log ("Detected PowerShell version: {0}" -f $PSVersionTable.PSVersion.ToString())

# =====================================================
# Resolve Environment IDs
# =====================================================
$ResolvedEnvironmentIds = @()

if ($EnvironmentId -and $EnvironmentIds) {
    Write-Log "Both -EnvironmentId and -EnvironmentIds were provided. Using -EnvironmentIds." -IsWarning
    $ResolvedEnvironmentIds = $EnvironmentIds
} elseif ($EnvironmentIds) {
    $ResolvedEnvironmentIds = $EnvironmentIds
} elseif ($EnvironmentId) {
    $ResolvedEnvironmentIds = @($EnvironmentId)
} else {
    Write-Log "No environment ID provided. Please supply -EnvironmentId or -EnvironmentIds." -IsError
    Write-Host "`nUsage example:" -ForegroundColor Yellow
    Write-Host "  .\get_nVectorMetrics.ps1 -LEApiToken `"token`" -EnvironmentId `"your-env-uuid`" -BaseUrl `"https://my.le.com`"`n" -ForegroundColor Yellow
    exit 1
}

Write-Log ("Resolved {0} environment ID(s) to query." -f $ResolvedEnvironmentIds.Count)

# =====================================================
# Resolve Time Range
# =====================================================
if ($StartTime -and $EndTime) {
    Write-Log ("Using provided time range: {0} to {1}" -f $StartTime, $EndTime)
} else {
    $EndTime   = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    $StartTime = (Get-Date).AddHours(-$LastHours).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    Write-Log ("Using last {0} hour(s). Range: {1} to {2}" -f $LastHours, $StartTime, $EndTime)
}

Write-Host ("Time range : {0} to {1}" -f $StartTime, $EndTime) -ForegroundColor Cyan
Write-Host ("Base URL   : {0}" -f $BaseUrl) -ForegroundColor Cyan
Write-Host ("API version: {0}" -f $ApiVersion) -ForegroundColor Cyan
Write-Host ("Output dir : {0}`n" -f $OutputDir) -ForegroundColor Cyan

# =====================================================
# TLS
# =====================================================
try {
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        Write-Log "Forced TLS 1.2 (PS5)."
    }
} catch {
    Write-Log ("Could not set TLS 1.2: {0}" -f $_.Exception.Message) -IsWarning
}

# =====================================================
# Certificate Functions
# =====================================================
function Get-RemoteCertificates {
    param([Parameter(Mandatory=$true)][string]$ServerHost, [int]$ServerPort = 443)
    $certList = New-Object System.Collections.ArrayList
    function Add-ChainFromLeaf([System.Security.Cryptography.X509Certificates.X509Certificate2]$leaf) {
        $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.RevocationMode = [System.Security.Cryptography.X509Certificates.X509RevocationMode]::NoCheck
        $null = $chain.Build($leaf)
        foreach ($elem in $chain.ChainElements) {
            $certObj = if ($elem.Certificate -is [System.Security.Cryptography.X509Certificates.X509Certificate2]) {
                $elem.Certificate
            } else {
                New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($elem.Certificate)
            }
            [void]$certList.Add($certObj)
        }
    }
    # Primary: SslStream
    try {
        $client = $null; $ssl = $null
        $client = New-Object System.Net.Sockets.TcpClient
        $client.Connect($ServerHost, $ServerPort)
        $stream = $client.GetStream()
        $ssl = New-Object System.Net.Security.SslStream($stream, $false, ({ $true }))
        $ssl.AuthenticateAsClient($ServerHost)
        $remote = $ssl.RemoteCertificate
        if ($remote) {
            $leaf = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($remote)
            Add-ChainFromLeaf -leaf $leaf
            try { $ssl.Close() } catch {}
            try { $client.Close() } catch {}
            return ,($certList.ToArray())
        }
    } catch {
        try { if ($ssl)    { $ssl.Dispose() }  } catch {}
        try { if ($client) { $client.Close() } } catch {}
    }
    # Fallback: HttpWebRequest / ServicePoint
    try {
        $oldCallback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { param($s,$c,$ch,$e) return $true }
        try {
            $req = [System.Net.HttpWebRequest]::Create("https://$ServerHost/")
            $req.Method = "HEAD"; $req.Timeout = 15000
            try {
                $resp = $req.GetResponse()
                try { $resp.Close() } catch {}
            } catch [System.Net.WebException] {
                if ($_.Exception.Response) { try { $_.Exception.Response.Close() } catch {} }
            }
            $svcCert = $req.ServicePoint.Certificate
            if ($svcCert) {
                $leaf2 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($svcCert)
                Add-ChainFromLeaf -leaf $leaf2
            } else { throw "No certificate available from ServicePoint for $ServerHost" }
        } finally {
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $oldCallback
        }
    } catch { throw $_ }
    return ,($certList.ToArray())
}

function Import-ServerCertificates {
    param([Parameter(Mandatory=$true)][string]$ServerHost, [int]$ServerPort = 443)
    Write-Log ("Fetching certificate from {0}:{1}..." -f $ServerHost, $ServerPort)
    $importedThumbs = @()
    try { $certs = Get-RemoteCertificates -ServerHost $ServerHost -ServerPort $ServerPort }
    catch {
        Write-Log ("Failed to obtain certificate: {0}" -f $_.Exception.Message) -IsError
        return ,@()
    }
    if (-not $certs -or $certs.Length -eq 0) {
        Write-Log "No certificates found." -IsError
        return ,@()
    }
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","CurrentUser")
    try {
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        foreach ($c in $certs) {
            $x2 = if ($c -is [System.Security.Cryptography.X509Certificates.X509Certificate2]) { $c } else {
                New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($c)
            }
            $thumb = $x2.Thumbprint
            $exists = $false
            foreach ($ec in $store.Certificates) { if ($ec.Thumbprint -eq $thumb) { $exists = $true; break } }
            if (-not $exists) {
                $store.Add($x2)
                $importedThumbs += $thumb
                Write-Log ("Imported cert thumbprint {0} into CurrentUser\Root" -f $thumb)
            } else {
                Write-Log ("Cert thumbprint {0} already present in CurrentUser\Root" -f $thumb)
            }
        }
    } catch {
        Write-Log ("Import error: {0}" -f $_.Exception.Message) -IsError
    } finally {
        try { $store.Close() } catch {}
    }
    return ,$importedThumbs
}

function Remove-ImportedCertificates {
    param([string[]]$Thumbprints)
    if (-not $Thumbprints -or $Thumbprints.Length -eq 0) { return }
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","CurrentUser")
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    try {
        foreach ($thumb in $Thumbprints) {
            $match = $store.Certificates | Where-Object { $_.Thumbprint -eq $thumb }
            if ($match) {
                $store.Remove($match)
                Write-Log ("Removed cert thumbprint {0} from CurrentUser\Root" -f $thumb)
            }
        }
    } finally { $store.Close() }
}

# =====================================================
# Main
# =====================================================
try {

    # Import cert if requested
    if ($ImportServerCert) {
        try {
            $leUri = [uri]$BaseUrl
            $lePort = if ($leUri.Port -ne -1 -and $leUri.Port -ne 0) { $leUri.Port } else { 443 }
            Write-Log "Importing Login Enterprise appliance certificate..."
            $script:ImportedCertThumbs = Import-ServerCertificates -ServerHost $leUri.Host -ServerPort $lePort
            if ($script:ImportedCertThumbs.Length -gt 0) {
                Write-Log ("Imported {0} certificate(s)." -f $script:ImportedCertThumbs.Length)
            }
        } catch {
            Write-Log ("Certificate import failed: {0}" -f $_.Exception.Message) -IsWarning
        }
    }

    $AllResults  = @()
    $AllDataRows = @()

    # Query each environment ID
    foreach ($envId in $ResolvedEnvironmentIds) {
        Write-Host ("`nQuerying environment: {0}" -f $envId) -ForegroundColor Yellow
        Write-Log ("Querying environment ID: {0}" -f $envId)

        # Build URL safely with UriBuilder
        try {
            $ub = New-Object System.UriBuilder($BaseUrl.TrimEnd("/"))
            $ub.Path = ($ub.Path.TrimEnd("/") + "/publicApi/$ApiVersion/platform-metrics").TrimStart("/")
            $queryParts = @(
                "from=$([uri]::EscapeDataString($StartTime))",
                "to=$([uri]::EscapeDataString($EndTime))",
                "environmentIds=$([uri]::EscapeDataString($envId))"
            )
            if ($MetricGroups) {
                foreach ($g in $MetricGroups) { $queryParts += "metricGroups=$([uri]::EscapeDataString($g))" }
            }
            $ub.Query = $queryParts -join "&"
            $FullUrl = $ub.Uri.AbsoluteUri
            Write-Log ("Constructed URL: {0}" -f $FullUrl)
        } catch {
            Write-Log ("Failed to construct URL for environment {0}: {1}" -f $envId, $_.Exception.Message) -IsError
            continue
        }

        $headers = @{
            "Authorization" = "Bearer $LEApiToken"
            "Accept"        = "application/json"
        }

        # Perform request
        $jsonResult = $null
        try {
            if ($PSVersionTable.PSVersion.Major -ge 7) {
                Write-Log "Using Invoke-RestMethod with -SkipCertificateCheck (PS7)."
                $jsonResult = Invoke-RestMethod -Uri $FullUrl -Method GET -Headers $headers -SkipCertificateCheck -ErrorAction Stop
            } else {
                Write-Log "Using HttpWebRequest (PS5)."
                $request = [System.Net.HttpWebRequest]::Create($FullUrl)
                $request.Method = "GET"
                $request.Headers.Add("Authorization", "Bearer $LEApiToken")
                $request.Accept = "application/json"
                $request.Timeout = 60000
                $response = $request.GetResponse()
                $stream   = $response.GetResponseStream()
                $reader   = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
                $rawJson  = $reader.ReadToEnd()
                $reader.Close(); $response.Close()
                $jsonResult = $rawJson | ConvertFrom-Json
            }
            Write-Log ("GET succeeded for environment {0}." -f $envId)
        } catch {
            Write-Log ("GET failed for environment {0}: {1}" -f $envId, $_.Exception.Message) -IsError
            continue
        }

        if ($jsonResult) {
            $AllResults += $jsonResult
            $seriesCount = 0
            foreach ($metric in $jsonResult) {
                $seriesCount++
                if ($metric.dataPoints) {
                    foreach ($dp in $metric.dataPoints) {
                        $AllDataRows += [PSCustomObject]@{
                            timestamp       = [string]$dp.timestamp
                            value           = $dp.value
                            metricId        = $metric.metricId
                            environmentKey  = $metric.environmentKey
                            displayName     = $metric.displayName
                            unit            = $metric.unit
                            instance        = $metric.instance
                            group           = $metric.group
                            componentType   = $metric.componentType
                        }
                    }
                }
            }
            Write-Host ("  Retrieved {0} metric series" -f $seriesCount) -ForegroundColor Green
            Write-Log ("Retrieved {0} metric series for environment {1}." -f $seriesCount, $envId)
        } else {
            Write-Log ("No data returned for environment {0}." -f $envId) -IsWarning
        }
    }

    # =====================================================
    # Summary
    # =====================================================
    Write-Host "`n========================================================================" -ForegroundColor Cyan
    Write-Host "  SUMMARY" -ForegroundColor Cyan
    Write-Host "========================================================================" -ForegroundColor Cyan

    if ($AllDataRows.Count -gt 0) {
        $AllDataRows | Group-Object -Property metricId | ForEach-Object {
            $sample = $_.Group | Select-Object -First 1
            Write-Host ("  {0} [{1}] — {2} data points" -f $sample.displayName, $sample.unit, $_.Count) -ForegroundColor White
        }
    } else {
        Write-Host "  No metrics found for the specified time range and environment(s)." -ForegroundColor Yellow
    }

    Write-Host ("`n  Total data points : {0}" -f $AllDataRows.Count) -ForegroundColor Cyan
    Write-Host ("  Environments queried: {0}" -f $ResolvedEnvironmentIds.Count) -ForegroundColor Cyan

    # =====================================================
    # Save Outputs
    # =====================================================
    if ($AllResults.Count -gt 0) {
        try {
            $AllResults | ConvertTo-Json -Depth 10 | Out-File $JsonPath -Encoding UTF8
            Write-Host ("`n  JSON saved : {0}" -f $JsonPath) -ForegroundColor Green
            Write-Log ("JSON saved to: {0}" -f $JsonPath)
        } catch {
            Write-Log ("Failed to write JSON: {0}" -f $_.Exception.Message) -IsError
        }
    }

    if ($AllDataRows.Count -gt 0) {
        try {
            $AllDataRows | Export-Csv -NoTypeInformation -Path $CsvPath -Encoding UTF8
            Write-Host ("  CSV saved  : {0}" -f $CsvPath) -ForegroundColor Green
            Write-Log ("CSV saved to: {0}" -f $CsvPath)
        } catch {
            Write-Log ("Failed to write CSV: {0}" -f $_.Exception.Message) -IsError
        }
    }

    Write-Host "`n========================================================================`n" -ForegroundColor Cyan
    Write-Log "Script completed successfully."

} finally {
    # Clean up imported certs unless -KeepCert was specified
    if ($ImportServerCert -and -not $KeepCert -and $script:ImportedCertThumbs.Length -gt 0) {
        try {
            Write-Log "Removing imported certificates..."
            Remove-ImportedCertificates -Thumbprints $script:ImportedCertThumbs
        } catch {
            Write-Log ("Failed to remove imported cert(s): {0}" -f $_.Exception.Message) -IsWarning
        }
    }
}