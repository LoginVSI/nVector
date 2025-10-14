param(
    [Parameter(Mandatory = $true)][string]$StartTime,
    [Parameter(Mandatory = $true)][string]$EndTime,
    [Parameter(Mandatory = $true)][string]$EnvironmentId,
    [string]$ApiAccessToken,
    [string]$BaseUrl,
    [string]$OutputCsvFilePath,
    [string]$OutputJsonFilePath,
    [string]$LogFilePath,
    [string]$ApiVersion = 'v7-preview',
    [switch]$ImportServerCert,
    [switch]$KeepCert,
    [switch]$Help
)

# defaults
$DefaultCsvFilePath  = "C:\temp\get_nVectorMetrics.csv"
$DefaultJsonFilePath = "C:\temp\get_nVectorMetrics.json"
$DefaultLogFilePath  = "C:\temp\get_nVectorMetrics_Log.txt"

if (-not $OutputCsvFilePath)  { $OutputCsvFilePath  = $DefaultCsvFilePath }
if (-not $OutputJsonFilePath) { $OutputJsonFilePath = $DefaultJsonFilePath }
if (-not $LogFilePath)        { $LogFilePath        = $DefaultLogFilePath }
$ScriptLogFile = $LogFilePath

# safer Write-Log (non-terminating)
function Write-Log {
    param([string]$Message, [switch]$IsError)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formatted = "$timestamp - $Message"
    try {
        $dir = Split-Path -Parent $ScriptLogFile
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        Add-Content -Path $ScriptLogFile -Value $formatted
    } catch {
        # swallow
        try { Write-Host ("WARNING: failed to write to log file: {0}" -f $_.Exception.Message) -ForegroundColor Yellow } catch {}
    }
    if ($IsError) { Write-Host $formatted -ForegroundColor Red } else { Write-Host $formatted }
}

Write-Log "==== Script invoked. ===="

if ($Help) {
    Write-Host "Usage: .\get_nVectorMetrics.ps1 -StartTime <ISO8601Z> -EndTime <ISO8601Z> -EnvironmentId <ID> [-ApiVersion <v7-preview>] [-ImportServerCert] [-KeepCert]"
    return
}

# detect PS version
$psVersion = $PSVersionTable.PSVersion
Write-Log ("Detected PowerShell version: {0}" -f $psVersion.ToString())

# defaults for token/baseurl if not provided
$DefaultApiAccessToken = "YOUR-DEFAULT-TOKEN-GOES-HERE"
$DefaultBaseUrl        = "https://myDomain.LoginEnterprise.com"

if ($ApiAccessToken) {
    $UsedApiAccessToken = $ApiAccessToken
    Write-Log "Using user-provided API token."
} else {
    $UsedApiAccessToken = $DefaultApiAccessToken
    Write-Log "No -ApiAccessToken provided; using default token."
}
if ($BaseUrl) {
    $UsedBaseUrl = $BaseUrl.TrimEnd('/')
    Write-Log ("Using user-provided BaseUrl: {0}" -f $UsedBaseUrl)
} else {
    $UsedBaseUrl = $DefaultBaseUrl.TrimEnd('/')
    Write-Log ("No -BaseUrl provided; using default: {0}" -f $UsedBaseUrl)
}

Write-Log ("Using ApiVersion: {0}" -f $ApiVersion)
Write-Log ("Using CSV path: {0}" -f $OutputCsvFilePath)
Write-Log ("Using JSON path: {0}" -f $OutputJsonFilePath)

# Force TLS1.2 for PS5
try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {}

# Build URL safely using UriBuilder
try {
    $ub = New-Object System.UriBuilder($UsedBaseUrl)
    $basePath = $ub.Path
    if ($basePath -eq $null) { $basePath = "" }
    $ub.Path = ($basePath.TrimEnd('/') + "/publicApi/$ApiVersion/platform-metrics").TrimStart('/')
    $ub.Query = "from=$([uri]::EscapeDataString($StartTime))&to=$([uri]::EscapeDataString($EndTime))&environmentIds=$([uri]::EscapeDataString($EnvironmentId))"
    $FullUrl = $ub.Uri.AbsoluteUri
} catch {
    Write-Log ("Failed to construct URL from BaseUrl '{0}': {1}" -f $UsedBaseUrl, $_.Exception.Message) -IsError
    return
}
Write-Log ("Constructed URL: {0}" -f $FullUrl)

# helper: fetch remote cert chain and import into CurrentUser\Root (returns list of thumbprints imported)
# Replace your existing Get-RemoteCertificates with this
function Get-RemoteCertificates {
    param(
        [Parameter(Mandatory=$true)][string]$ServerHost,
        [int]$ServerPort = 443
    )

    # Use ArrayList to avoid += overload issues with X509Certificate2 objects
    $certList = New-Object System.Collections.ArrayList

    # Helper to build chain and add elements to the ArrayList
    function Add-ChainFromLeaf([System.Security.Cryptography.X509Certificates.X509Certificate2]$leaf) {
        $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.RevocationMode = [System.Security.Cryptography.X509Certificates.X509RevocationMode]::NoCheck
        $null = $chain.Build($leaf)
        foreach ($elem in $chain.ChainElements) {
            # ensure X509Certificate2 in list
            $certObj = if ($elem.Certificate -is [System.Security.Cryptography.X509Certificates.X509Certificate2]) { $elem.Certificate } else { New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($elem.Certificate) }
            [void]$certList.Add($certObj)
        }
    }

    # First attempt: SslStream (best-effort)
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
        } else {
            throw "SslStream returned no remote certificate."
        }
    } catch {
        # cleanup and fall through to fallback approach
        try { if ($ssl) { $ssl.Dispose() } } catch {}
        try { if ($client) { $client.Close() } } catch {}
        # swallow and continue to fallback
    }

    # Fallback: HttpWebRequest -> ServicePoint.Certificate
    try {
        $oldCallback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { param($s,$c,$ch,$e) return $true }
        try {
            $uri = "https://$ServerHost/"
            $req = [System.Net.HttpWebRequest]::Create($uri)
            $req.Method = 'HEAD'
            $req.Timeout = 15000
            try {
                $resp = $req.GetResponse()
                try { $resp.Close() } catch {}
            } catch [System.Net.WebException] {
                # HEAD may be rejected; continue — ServicePoint.Certificate might still be populated.
                if ($_.Exception.Response) { try { $_.Exception.Response.Close() } catch {} }
            }
            $svcCert = $req.ServicePoint.Certificate
            if ($svcCert) {
                $leaf2 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($svcCert)
                Add-ChainFromLeaf -leaf $leaf2
            } else {
                throw "No certificate available from ServicePoint for $ServerHost"
            }
        } finally {
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $oldCallback
        }
    } catch {
        throw $_
    }

    return ,($certList.ToArray())
}

function Import-ServerCertificates {
    param(
        [Parameter(Mandatory=$true)][string]$ServerHost,
        [int]$ServerPort = 443,
        [switch]$Keep
    )

    Write-Log ("Attempting to fetch and import cert(s) from {0}:{1}" -f $ServerHost,$ServerPort)
    $importedThumbs = @()

    try {
        $certs = Get-RemoteCertificates -ServerHost $ServerHost -ServerPort $ServerPort
    } catch {
        Write-Log ("Failed to obtain remote certificate from {0}:{1}: {2}" -f $ServerHost,$ServerPort,$_.Exception.Message) -IsError
        return ,@()
    }

    if (-not $certs -or $certs.Length -eq 0) {
        Write-Log ("No certificates found for {0}:{1}" -f $ServerHost,$ServerPort) -IsError
        return ,@()
    }

    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","CurrentUser")
    try {
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        foreach ($c in $certs) {
            # ensure X509Certificate2
            $x2 = if ($c -is [System.Security.Cryptography.X509Certificates.X509Certificate2]) { $c } else { New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($c) }
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

    if ($importedThumbs.Count -eq 0) {
        Write-Log ("Imported 0 cert(s)")
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

# optionally import server certs
$importedThumbs = @()
if ($ImportServerCert) {
    try {
        $u = [uri]$UsedBaseUrl
        $port = 443
        if ($u.Port -ne -1 -and $u.Port -ne 0) { $port = $u.Port }
        $importedThumbs = Import-ServerCertificates -ServerHost $u.Host -ServerPort $port -Keep:$KeepCert
    } catch {
        Write-Log ("ImportServerCert failed: {0}" -f $_.Exception.Message) -IsError
    }
}

# Prepare headers
$Headers = @{
    "Authorization" = "Bearer $UsedApiAccessToken"
    "Accept"        = "application/json"
}

# perform request
$jsonString = $null
$JsonResponse = $null

if ($psVersion.Major -ge 7) {
    Write-Log "Using Invoke-RestMethod with -SkipCertificateCheck (PowerShell 7.x)."
    try {
        $jsonResult = Invoke-RestMethod -Uri $FullUrl -Method GET -Headers $Headers -SkipCertificateCheck -ErrorAction Stop
        Write-Log "GET request succeeded."
        if ($jsonResult -is [string]) { $jsonString = $jsonResult } else { $jsonString = $jsonResult | ConvertTo-Json -Depth 10 }
        $JsonResponse = if ($jsonResult -is [string]) { $jsonResult | ConvertFrom-Json } else { $jsonResult }
    } catch {
        Write-Log ("Error during GET request: {0}" -f $_.Exception.Message) -IsError
    }
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
        $reader = New-Object System.IO.StreamReader($stream,[System.Text.Encoding]::UTF8)
        $jsonString = $reader.ReadToEnd()
        $reader.Close()
        $response.Close()
        Write-Log "GET request succeeded."
        try {
            $JsonResponse = $jsonString | ConvertFrom-Json
        } catch {
            Write-Log ("Failed to parse JSON response: {0}" -f $_.Exception.Message) -IsError
        }
    } catch {
        Write-Log ("Error during GET request: {0}" -f $_.Exception.Message) -IsError
    }
}

# cleanup imported certs when requested (only if we imported and KeepCert not set)
if ($ImportServerCert -and -not $KeepCert) {
    try {
        Remove-ImportedCertificates -Thumbprints $importedThumbs
    } catch {
        Write-Log ("Failed to remove imported cert(s): {0}" -f $_.Exception.Message) -IsError
    }
}

# Save raw JSON if present
if ($null -ne $jsonString) {
    try {
        $jsonString | Out-File -FilePath $OutputJsonFilePath -Encoding UTF8
        Write-Log ("JSON saved to: {0}" -f $OutputJsonFilePath)
    } catch {
        Write-Log ("Failed to write JSON output: {0}" -f $_.Exception.Message) -IsError
    }
} else {
    Write-Log "No JSON response to save."
}

# Convert to CSV
Write-Log "Converting JSON to CSV..."
$AllDataRows = @()
if ($JsonResponse) {
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
        try {
            $AllDataRows | Export-Csv -NoTypeInformation -Path $OutputCsvFilePath -Encoding UTF8
            Write-Log ("CSV saved to: {0}" -f $OutputCsvFilePath)
        } catch {
            Write-Log ("Failed to write CSV output: {0}" -f $_.Exception.Message) -IsError
        }
    }
} else {
    Write-Log "No parsed JSON to convert to CSV."
}

Write-Log "Script completed."
