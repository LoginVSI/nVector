<#
Create-NvidiaGpuSessionMetricGroup.ps1
version 1.0.0

PREVIEW / INTERIM RELEASE
This script is part of the Login Enterprise + NVIDIA nVector integration
preview. Refer to github.com/LoginVSI/nVector and docs.loginvsi.com
for documentation.

PURPOSE:
  One-time setup script. Run this once on (or targeting) a machine that
  has an NVIDIA enterprise GPU with PerfMon counters available.

  The script will:
    1. Discover the NVIDIA GPU PerfMon instance name dynamically
    2. Verify the five required GPU counters are present
    3. Create five NVIDIA GPU session metric definitions in Login Enterprise
    4. Fetch the built-in CPU (metricId=1) and Memory (metricId=2) metric keys
    5. Create a new session metric group containing all seven metrics
    6. Output a full summary including group ID and next steps

  After running, assign the created group to your Continuous or Load Test:
  LE UI > Configuration > Manage Tests > [Your Test] > Test Settings
  > Session Metrics > Metrics Group

  Verify the result in the LE UI under:
  Configuration > Session Metrics > Groups

NOTES:
  - Run this once per target image / GPU profile. Re-running will create
    duplicate metric definitions. If you need to re-run, delete the
    previously created definitions first via the LE UI or API.
  - Session metric definitions cannot be edited after creation in LE.
    If you need to change a definition, delete and recreate it.
  - All seven metrics (5x GPU + CPU + Memory) share the same unit (%)
    and will chart together on the same y-axis in LE session metric charts.
  - The built-in CPU and Memory metric keys are appliance-specific GUIDs.
    This script discovers them at runtime by metricId (1=CPU, 2=Memory)
    rather than hardcoding keys.

USAGE EXAMPLES:
  # Basic — uses defaults in script body:
  .\Create-NvidiaGpuSessionMetricGroup.ps1

  # Override appliance URL and token:
  .\Create-NvidiaGpuSessionMetricGroup.ps1 -BaseUrl "https://my-le.example.com" -ConfigurationAccessToken "mytoken"

  # For appliances with self-signed certs (imports cert, removes after run):
  .\Create-NvidiaGpuSessionMetricGroup.ps1 -BaseUrl "https://my-le.example.com" -ConfigurationAccessToken "mytoken" -ImportServerCert

  # Keep the imported cert after run:
  .\Create-NvidiaGpuSessionMetricGroup.ps1 -ImportServerCert -KeepCert

  # Target a remote GPU machine for counter discovery:
  .\Create-NvidiaGpuSessionMetricGroup.ps1 -TargetComputer "my-gpu-vm"

  # Custom group name:
  .\Create-NvidiaGpuSessionMetricGroup.ps1 -GroupName "NVIDIA A10-4Q GPU Metrics"
#>

param(
    [string]$BaseUrl                  = "",
    [string]$ConfigurationAccessToken = "",
    [string]$GroupName                = "",   # auto-set from GPU model if left empty
    [string]$GroupDescription         = "",
    [string]$TargetComputer           = "",   # leave empty to query local machine
    [switch]$ImportServerCert,               # imports the LE appliance SSL cert into CurrentUser\Root (use for self-signed certs)
    [switch]$KeepCert                        # keeps imported cert after run (default: removed on exit)
)

$ScriptVersion = "1.0.0"

# =====================================================
# Config defaults — edit these to match your environment
# CLI params above override at runtime
# =====================================================
if ($BaseUrl                  -eq "") { $BaseUrl                  = "https://myDomain.LoginEnterprise.com/" }
if ($ConfigurationAccessToken -eq "") { $ConfigurationAccessToken = "abcd1234abcd1234abcd1234abcd1234abcd1234abc" }
if ($GroupDescription         -eq "") { $GroupDescription         = "NVIDIA GPU session metrics — auto-created by Create-NvidiaGpuSessionMetricGroup.ps1 v$ScriptVersion" }

# API version
$ApiVersion = "v8-preview"

# Tag applied to all created GPU metrics
$GpuTag = "gpu"

# Retry settings
$ApiRetryCount   = 3
$ApiRetryDelayMs = 2000

# Built-in metric IDs — these are fixed across all LE appliances
# metricId 1 = CPU utilization (display name: "CPU", unit: %)
# metricId 2 = Memory utilization (display name: "Memory", unit: %)
$BuiltInCpuMetricId    = 1
$BuiltInMemoryMetricId = 2

# =====================================================
# Logging — timestamped log saved to %TEMP%
# =====================================================
$Timestamp = (Get-Date).ToString('yyyyMMddTHHmmss')
$LogFile   = Join-Path $env:TEMP "${Timestamp}_Create-NvidiaGpuSessionMetricGroup.log"

$ErrorActionPreference = 'Continue'

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss"), $Level, $Message
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -ErrorAction SilentlyContinue
}

Write-Log "============================================================"
Write-Log "Create-NvidiaGpuSessionMetricGroup.ps1 v$ScriptVersion"
Write-Log "PREVIEW / INTERIM RELEASE — Login VSI + NVIDIA nVector"
Write-Log "============================================================"
Write-Log ("BaseUrl:        {0}" -f $BaseUrl)
$targetDisplay = if ($TargetComputer -ne '') { $TargetComputer } else { "(local)" }
Write-Log ("TargetComputer: {0}" -f $targetDisplay)
Write-Log ("Log file:       {0}" -f $LogFile)

$script:ImportedCertThumbs = @()

# =====================================================
# TLS
# =====================================================
try {
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        Write-Log "Forced TLS 1.2 (PS5)."
    }
} catch {
    Write-Log ("Could not set TLS 1.2: {0}" -f $_.Exception.Message) "WARN"
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
        Write-Log ("Failed to obtain certificate: {0}" -f $_.Exception.Message) "WARN"
        return ,@()
    }
    if (-not $certs -or $certs.Length -eq 0) {
        Write-Log "No certificates found." "WARN"
        return ,@()
    }
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","CurrentUser")
    try {
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        foreach ($c in $certs) {
            $x2 = if ($c -is [System.Security.Cryptography.X509Certificates.X509Certificate2]) { $c } else {
                New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($c)
            }
            $thumb  = $x2.Thumbprint
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
        Write-Log ("Import error: {0}" -f $_.Exception.Message) "WARN"
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

# Import cert if requested
if ($ImportServerCert) {
    try {
        $leUri  = [uri]$BaseUrl
        $lePort = if ($leUri.Port -ne -1 -and $leUri.Port -ne 0) { $leUri.Port } else { 443 }
        Write-Log "Importing Login Enterprise appliance certificate..."
        $script:ImportedCertThumbs = Import-ServerCertificates -ServerHost $leUri.Host -ServerPort $lePort
        if ($script:ImportedCertThumbs.Length -gt 0) {
            Write-Log ("Imported {0} certificate(s)." -f $script:ImportedCertThumbs.Length)
        }
        # PS5 Invoke-WebRequest does not honour certs imported into CurrentUser\Root in the
        # same session — set bypass callback so the imported cert takes effect immediately.
        # Only set when -ImportServerCert is used (i.e. appliance has a self-signed cert).
        # Appliances with CA-signed certs do not need this and should not use -ImportServerCert.
        if ($PSVersionTable.PSVersion.Major -lt 7) {
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
            Write-Log "SSL validation bypass set for self-signed cert (PS5)."
        }
    } catch {
        Write-Log ("Certificate import failed: {0}" -f $_.Exception.Message) "WARN"
    }
}

# =====================================================
# API helper with retry
# =====================================================
function Invoke-LeApi {
    param(
        [string]$Method,
        [string]$EndpointPath,
        [object]$Body        = $null,
        [hashtable]$Query    = $null
    )

    $url = $BaseUrl.TrimEnd('/') + '/publicApi/' + $ApiVersion + '/' + $EndpointPath.TrimStart('/')

    # Append query string params if provided
    if ($Query -and $Query.Count -gt 0) {
        $qs = ($Query.GetEnumerator() | ForEach-Object { "{0}={1}" -f $_.Key, $_.Value }) -join "&"
        $url = $url + "?" + $qs
    }

    $hdr = @{
        Authorization  = "Bearer $ConfigurationAccessToken"
        "Content-Type" = "application/json"
        "Accept"       = "application/json"
    }

    Write-Log ("API --> {0} {1}" -f $Method, $url)

    $bodyJson = $null
    if ($null -ne $Body) {
        if ($Body -is [array]) {
            $bodyJson = ConvertTo-Json -InputObject @($Body) -Depth 10 -Compress
        } else {
            $bodyJson = $Body | ConvertTo-Json -Depth 10 -Compress
        }
        Write-Log ("Request body: {0}" -f $bodyJson)
    }

    for ($attempt = 1; $attempt -le $ApiRetryCount; $attempt++) {
        try {
            $params = @{
                Uri             = $url
                Method          = $Method
                Headers         = $hdr
                UseBasicParsing = $true
                ErrorAction     = "Stop"
            }
            if ($bodyJson) { $params.Body = $bodyJson; $params.ContentType = "application/json" }
            $webResponse = Invoke-WebRequest @params
            if ($webResponse.Content) {
                return $webResponse.Content | ConvertFrom-Json
            }
            return $true
        } catch {
            $status = $null
            $detail = $_.Exception.Message
            try { $status = $_.Exception.Response.StatusCode.value__ } catch {}
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $detail = $reader.ReadToEnd()
                $reader.Close()
            } catch {
                try { $detail = $_.ErrorDetails.Message } catch {}
            }
            Write-Log ("API {0} {1} failed (attempt {2}/{3}) — HTTP {4}: {5}" -f $Method, $EndpointPath, $attempt, $ApiRetryCount, $status, $detail) "WARN"
            if ($attempt -lt $ApiRetryCount) {
                Start-Sleep -Milliseconds $ApiRetryDelayMs
            }
        }
    }

    Write-Log ("ERROR: API call failed after {0} attempts: {1} {2}" -f $ApiRetryCount, $Method, $EndpointPath) "ERROR"
    return $null
}

# =====================================================
# Step 1 — Discover NVIDIA GPU PerfMon instance
# =====================================================
Write-Log "------------------------------------------------------------"
Write-Log "Step 1: Discovering NVIDIA GPU PerfMon instance..."

try {
    $getCounterParams = @{ ListSet = "NVIDIA GPU"; ErrorAction = "Stop" }
    if ($TargetComputer -ne "") { $getCounterParams.ComputerName = $TargetComputer }
    $counterSet = Get-Counter @getCounterParams
} catch {
    Write-Log ("ERROR: Could not query 'NVIDIA GPU' counter set. Ensure the NVIDIA driver is installed and GPU PerfMon counters are available on the target machine. Details: {0}" -f $_.Exception.Message) "ERROR"
    exit 1
}

$pathsWithInstances = $counterSet.PathsWithInstances
if (-not $pathsWithInstances -or $pathsWithInstances.Count -eq 0) {
    Write-Log "ERROR: No NVIDIA GPU instances found in PerfMon. Cannot continue." "ERROR"
    exit 1
}

# Parse instance name from first counter path
# Format: \NVIDIA GPU(#0 Tesla M60 (id=1, NVAPI ID=65536))\% GPU Usage
$firstPath    = $pathsWithInstances[0]
$instanceName = $null
if ($firstPath -match 'NVIDIA GPU\((.+)\)\\') {
    $instanceName = $Matches[1]
}

if (-not $instanceName) {
    Write-Log ("ERROR: Could not parse instance name from: {0}" -f $firstPath) "ERROR"
    exit 1
}

Write-Log ("GPU instance found: {0}" -f $instanceName)

# Extract GPU model and NVAPI ID for use in display names
# Instance format: #0 Tesla M60 (id=1, NVAPI ID=65536)
$gpuModel = $null
$nvapiId  = $null

if ($instanceName -match '#\d+\s+(.+?)\s+\(id=\d+,\s+NVAPI ID=(\d+)\)') {
    $gpuModel = $Matches[1]   # e.g. "Tesla M60" or "A10-4Q"
    $nvapiId  = $Matches[2]   # e.g. "65536" or "131072"
    Write-Log ("GPU model: {0}  |  NVAPI ID: {1}" -f $gpuModel, $nvapiId)
} else {
    Write-Log "WARNING: Could not parse GPU model/NVAPI ID from instance name. Using full instance string for display names." "WARN"
    $gpuModel = $instanceName
    $nvapiId  = "unknown"
}

# Auto-set group name if not passed in
if ($GroupName -eq "") {
    $GroupName = "NVIDIA GPU - {0} (NVAPI {1})" -f $gpuModel, $nvapiId
    Write-Log ("Group name auto-set: {0}" -f $GroupName)
}

# =====================================================
# Step 2 — Verify required counters are present
# =====================================================
Write-Log "------------------------------------------------------------"
Write-Log "Step 2: Verifying required GPU counters are present..."

$requiredCounters = @(
    "% GPU Usage",
    "% GPU Memory Usage",
    "% FB Usage",
    "% Video Decoder Usage",
    "% Video Encoder Usage"
)

$missingCounters = @()
foreach ($counter in $requiredCounters) {
    $match = $pathsWithInstances | Where-Object { $_ -match [regex]::Escape($counter) }
    if ($match) {
        Write-Log ("  [OK] {0}" -f $counter)
    } else {
        Write-Log ("  [MISSING] {0}" -f $counter) "WARN"
        $missingCounters += $counter
    }
}

if ($missingCounters.Count -gt 0) {
    Write-Log ("WARNING: {0} counter(s) not found on this machine. They will still be created in LE but may not collect data. Verify the NVIDIA driver supports these counters." -f $missingCounters.Count) "WARN"
}

# =====================================================
# Step 3 — Build and create GPU metric definitions in LE
# =====================================================
Write-Log "------------------------------------------------------------"
Write-Log "Step 3: Creating GPU session metric definitions in Login Enterprise..."

# Naming convention: "NVIDIA <model> (<NVAPI ID>) - <counter>"
# e.g. "NVIDIA Tesla M60 (65536) - % GPU Usage"
# Display name uses same format — identifiable in LE if multiple GPU profiles exist

function New-GpuMetricDefinition {
    param([string]$CounterName)
    return @{
        type        = "PerformanceCounter"
        name        = "NVIDIA {0} ({1}) - {2}" -f $gpuModel, $nvapiId, $CounterName
        description = "NVIDIA GPU PerfMon — {0} | Instance: {1}" -f $CounterName, $instanceName
        tag         = $GpuTag
        measurement = @{
            counterCategory = "NVIDIA GPU"
            counterName     = $CounterName
            counterInstance = $instanceName
            displayName     = "{0} ({1}) - {2}" -f $gpuModel, $nvapiId, $CounterName
            unit            = "%"
        }
    }
}

$gpuCountersToCreate = @(
    "% GPU Usage",
    "% GPU Memory Usage",
    "% FB Usage",
    "% Video Decoder Usage",
    "% Video Encoder Usage"
)

$createdMetricKeys = @()

foreach ($counterName in $gpuCountersToCreate) {
    $def = New-GpuMetricDefinition -CounterName $counterName
    Write-Log ("  Creating: {0}" -f $def.name)

    $result = Invoke-LeApi -Method "POST" -EndpointPath "user-session-metric-definitions" -Body $def

    if ($result -and $result.id) {
        Write-Log ("    [OK] ID: {0}" -f $result.id)
        $createdMetricKeys += $result.id
    } else {
        Write-Log ("    [ERROR] Failed to create: {0}" -f $def.name) "ERROR"
        Write-Log "    Aborting. Clean up any partially created definitions in LE UI (Configuration > Session Metrics > Metrics) before retrying." "ERROR"
        exit 1
    }
}

Write-Log ("{0} GPU metric definitions created successfully." -f $createdMetricKeys.Count)

# =====================================================
# Step 4 — Fetch built-in CPU and Memory metric keys
# Built-in metrics are identified by metricId (1=CPU, 2=Memory)
# Keys are appliance-specific GUIDs — must be discovered at runtime
# =====================================================
Write-Log "------------------------------------------------------------"
Write-Log "Step 4: Fetching built-in CPU and Memory metric keys..."

# Use count=100 to get all definitions; offset if needed
$allMetricsResponse = Invoke-LeApi -Method "GET" -EndpointPath "user-session-metric-definitions" `
    -Query @{ count = 100; offset = 0; includeTotalCount = "true" }

if (-not $allMetricsResponse -or -not $allMetricsResponse.items) {
    Write-Log "ERROR: Could not retrieve session metric definitions from LE." "ERROR"
    exit 1
}

$allMetrics = $allMetricsResponse.items
Write-Log ("{0} total metric definitions found in LE." -f $allMetrics.Count)

# Find built-ins by metricId — reliable across all appliances
$cpuMetric = $allMetrics | Where-Object { $_.type -eq "BuiltIn" -and $_.measurement.metricId -eq $BuiltInCpuMetricId }
$memMetric = $allMetrics | Where-Object { $_.type -eq "BuiltIn" -and $_.measurement.metricId -eq $BuiltInMemoryMetricId }

# Fallback: match by display name if metricId path doesn't resolve
if (-not $cpuMetric) {
    $cpuMetric = $allMetrics | Where-Object { $_.type -eq "BuiltIn" -and $_.measurement.displayName -match "^CPU" }
    if ($cpuMetric) { Write-Log "  (CPU found via displayName fallback)" "WARN" }
}
if (-not $memMetric) {
    $memMetric = $allMetrics | Where-Object { $_.type -eq "BuiltIn" -and $_.measurement.displayName -match "^Mem" }
    if ($memMetric) { Write-Log "  (Memory found via displayName fallback)" "WARN" }
}

$builtInKeys = @()

if ($cpuMetric) {
    Write-Log ("  [OK] Built-in CPU metric: displayName='{0}' unit='{1}' key={2}" -f $cpuMetric.measurement.displayName, $cpuMetric.measurement.unit, $cpuMetric.key)
    $builtInKeys += $cpuMetric.key
} else {
    Write-Log "  [WARN] Built-in CPU metric not found — it will not be added to the group." "WARN"
}

if ($memMetric) {
    $memMetric = $memMetric | Select-Object -First 1
    Write-Log ("  [OK] Built-in Memory metric: displayName='{0}' unit='{1}' key={2}" -f $memMetric.measurement.displayName, $memMetric.measurement.unit, $memMetric.key)
    $builtInKeys += $memMetric.key
} else {
    Write-Log "  [WARN] Built-in Memory metric not found — it will not be added to the group." "WARN"
}

# =====================================================
# Step 5 — Create the session metric group
# =====================================================
Write-Log "------------------------------------------------------------"
Write-Log ("Step 5: Creating session metric group '{0}'..." -f $GroupName)

# All keys: 5 GPU metrics + built-in CPU + built-in Memory
$allDefinitionKeys = $createdMetricKeys + $builtInKeys

Write-Log ("{0} total metrics will be in this group ({1} GPU + {2} built-in)." -f $allDefinitionKeys.Count, $createdMetricKeys.Count, $builtInKeys.Count)

$groupBody = @{
    name           = $GroupName
    description    = $GroupDescription
    definitionKeys = $allDefinitionKeys
}

$groupResult = Invoke-LeApi -Method "POST" -EndpointPath "session-metric-definition-groups" -Body $groupBody

if (-not $groupResult -or -not $groupResult.id) {
    Write-Log "ERROR: Failed to create session metric group." "ERROR"
    exit 1
}

$groupId = $groupResult.id
Write-Log ("[OK] Group created — ID: {0}" -f $groupId)

# =====================================================
# Step 6 — Verify: GET the group back
# =====================================================
Write-Log "------------------------------------------------------------"
Write-Log "Step 6: Verifying group via GET..."

$verifyGroup = Invoke-LeApi -Method "GET" -EndpointPath ("session-metric-definition-groups/{0}" -f $groupId)

if ($verifyGroup) {
    Write-Log ("[OK] Group verified: name='{0}' memberCount={1}" -f $verifyGroup.name, $verifyGroup.memberCount)
} else {
    Write-Log "WARNING: Could not verify group via GET. The group may still have been created — check LE UI." "WARN"
}

# =====================================================
# Summary output
# =====================================================
Write-Log "============================================================"
Write-Log "SUCCESS — SUMMARY"
Write-Log "============================================================"
Write-Log ("GPU instance:         {0}" -f $instanceName)
Write-Log ("GPU model:            {0}" -f $gpuModel)
Write-Log ("NVAPI ID:             {0}" -f $nvapiId)
Write-Log ("Group name:           {0}" -f $GroupName)
Write-Log ("Group ID:             {0}" -f $groupId)
Write-Log ("Total metrics in group: {0}" -f $allDefinitionKeys.Count)
Write-Log ""
Write-Log "GPU metrics created:"
for ($i = 0; $i -lt $gpuCountersToCreate.Count; $i++) {
    Write-Log ("  [{0}] {1} — key: {2}" -f ($i+1), $gpuCountersToCreate[$i], $createdMetricKeys[$i])
}
Write-Log ""
Write-Log "Built-in metrics added:"
if ($cpuMetric) { Write-Log ("  CPU     — key: {0}" -f $cpuMetric.key) }
if ($memMetric) { Write-Log ("  Memory  — key: {0}" -f $memMetric.key) }
Write-Log ""
Write-Log "------------------------------------------------------------"
Write-Log "NEXT STEPS:"
Write-Log "  1. Verify the group in LE UI:"
Write-Log "     Configuration > Session Metrics > Groups"
Write-Log ("     Look for: '{0}'" -f $GroupName)
Write-Log "  2. Assign the group to your test:"
Write-Log "     Configuration > Manage Tests > [Your Test]"
Write-Log "     > Test Settings > Session Metrics > Metrics Group"
Write-Log ("     Group to select: '{0}'" -f $GroupName)
Write-Log "  3. Run a short no-action test to validate all metrics are collecting."
Write-Log "     (LE UI > run test > view results > Session Metrics tab)"
Write-Log "  4. All 7 metrics share unit '%' and will chart together on one y-axis."
Write-Log "============================================================"
Write-Log ("Log saved to: {0}" -f $LogFile)

# Cert cleanup
if ($ImportServerCert -and -not $KeepCert -and $script:ImportedCertThumbs.Length -gt 0) {
    try { Remove-ImportedCertificates -Thumbprints $script:ImportedCertThumbs }
    catch { Write-Log ("Cert cleanup failed: {0}" -f $_.Exception.Message) "WARN" }
}
