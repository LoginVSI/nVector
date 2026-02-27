<#
Add-SessionMetrics.ps1
version 1.0.0

PURPOSE:
  A generic, interactive script for discovering Windows PerfMon and WMI
  performance counters on any Windows machine, then creating session metric
  definitions in a Login Enterprise appliance via the public API.

  Optionally adds created metrics to an existing session metric group,
  creates a new group, or leaves them unassigned for later organization.

  This script has no dependency on any specific vendor, protocol, or
  workload. It can be used to register any PerfMon counter or WMI class
  property as a Login Enterprise session metric.

RUN ON:
  Any Windows machine that:
    a) Has the PerfMon/WMI counters you want to discover installed/available
    b) Can reach the Login Enterprise appliance API over the network
  This is NOT required to be a launcher, target, or any specific machine role.

TYPICAL USE CASES:
  - Finding and registering remoting protocol counters (RemoteFX, Citrix,
    Omnissa Blast, PCoIP) once the correct counter name is known
  - Registering any custom PerfMon or WMI counter as a session metric
  - Bulk-adding metrics without using the LE Swagger UI
  - Exploring what performance counters are available on a given machine

NOTES:
  - Session metric definitions cannot be edited after creation in LE.
    If you need to change a definition, delete it in LE UI and re-run.
  - Re-running with the same search string will create duplicate definitions.
    Check existing metrics first: LE UI > Configuration > Session Metrics.
  - WMI property units vary widely (%, ms, Bytes, Frames/sec, Count, MB, etc.)
    The script will suggest a unit per property based on the property name and
    prompt you to confirm or override. You can always correct units in LE UI
    by deleting and recreating the metric definition.
  - PerfMon: counterInstance is set from the discovered instance. If a counter
    has multiple instances you may want to run again and select specifically.

USAGE EXAMPLES:
  # Fully interactive (prompted for everything):
  .\Add-SessionMetrics.ps1

  # Provide search string upfront, still interactive for selection:
  .\Add-SessionMetrics.ps1 -SearchString "RemoteFX"

  # Override LE appliance connection:
  .\Add-SessionMetrics.ps1 -BaseUrl "https://my-le.example.com" -ConfigurationAccessToken "mytoken"

  # For appliances with self-signed certs (imports cert, removes after run):
  .\Add-SessionMetrics.ps1 -BaseUrl "https://my-le.example.com" -ConfigurationAccessToken "mytoken" -ImportServerCert

  # Keep the imported cert after run (useful if running repeatedly):
  .\Add-SessionMetrics.ps1 -ImportServerCert -KeepCert

  # Non-interactive group assignment (add to existing group by ID):
  .\Add-SessionMetrics.ps1 -ExistingGroupId "a2f7d6d2-1dd8-4f42-a455-b741535527e1"

  # Non-interactive group creation:
  .\Add-SessionMetrics.ps1 -NewGroupName "Blast Protocol Metrics"
#>

param(
    [string]$BaseUrl                  = "",
    [string]$ConfigurationAccessToken = "",
    [string]$SearchString             = "",   # counter/class search string; prompted interactively if not provided
    [string]$MetricTag                = "",   # optional tag applied to all created metrics
    [string]$ExistingGroupId          = "",   # if provided, adds created metrics to this group (skips group prompt)
    [string]$NewGroupName             = "",   # if provided, creates a new group with this name (skips group prompt)
    [string]$NewGroupDescription      = "",
    [switch]$ImportServerCert,               # imports the LE appliance SSL cert into CurrentUser\Root (use for self-signed certs)
    [switch]$KeepCert                        # keeps imported cert after run (default: removed on exit)
)

$ScriptVersion = "1.0.0"

# =====================================================
# Config defaults — edit these for your environment
# All can also be overridden via CLI params above
# =====================================================
if ($BaseUrl                  -eq "") { $BaseUrl                  = "https://myDomain.LoginEnterprise.com/" }
if ($ConfigurationAccessToken -eq "") { $ConfigurationAccessToken = "your-api-token-here" }

$ApiVersion      = "v8-preview"
$ApiRetryCount   = 3
$ApiRetryDelayMs = 2000

# WMI namespace to search
$WmiNamespace = "root\cimv2"

# WMI system/meta properties to exclude from property enumeration
$WmiSystemProperties = @(
    "Caption", "Description", "Name", "Frequency_Object", "Frequency_PerfTime",
    "Frequency_Sys100NS", "Timestamp_Object", "Timestamp_PerfTime", "Timestamp_Sys100NS",
    "__CLASS", "__DYNASTY", "__GENUS", "__NAMESPACE", "__PATH", "__PROPERTY_COUNT",
    "__RELPATH", "__SERVER", "__SUPERCLASS", "__DERIVATION"
)

# =====================================================
# Logging — timestamped log to %TEMP%
# =====================================================
$Timestamp = (Get-Date).ToString('yyyyMMddTHHmmss')
$LogFile   = Join-Path $env:TEMP "${Timestamp}_Add-SessionMetrics.log"
$ErrorActionPreference = 'Continue'

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss"), $Level, $Message
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -ErrorAction SilentlyContinue
}

function Write-Console {
    param([string]$Message = "", [ConsoleColor]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

Write-Log "============================================================"
Write-Log "Add-SessionMetrics.ps1 v$ScriptVersion"
Write-Log "Login Enterprise — Generic Session Metric Discovery and Registration"
Write-Log "============================================================"
Write-Log ("BaseUrl:  {0}" -f $BaseUrl)
Write-Log ("Machine:  {0}" -f $env:COMPUTERNAME)
Write-Log ("Log file: {0}" -f $LogFile)

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
        [object]$Body     = $null,
        [hashtable]$Query = $null
    )

    $url = $BaseUrl.TrimEnd('/') + '/publicApi/' + $ApiVersion + '/' + $EndpointPath.TrimStart('/')
    if ($Query -and $Query.Count -gt 0) {
        $amp = [char]38
        $qs  = ($Query.GetEnumerator() | ForEach-Object { "{0}={1}" -f $_.Key, $_.Value }) -join $amp
        $url = $url + "?" + $qs
    }

    Write-Log ("API --> {0} {1}" -f $Method, $url) "INFO"

    $hdr = @{
        Authorization  = "Bearer $ConfigurationAccessToken"
        "Content-Type" = "application/json"
        "Accept"       = "application/json"
    }

    $bodyJson = $null
    if ($null -ne $Body) {
        if ($Body -is [array]) {
            $bodyJson = ConvertTo-Json -InputObject @($Body) -Depth 10 -Compress
        } else {
            $bodyJson = $Body | ConvertTo-Json -Depth 10 -Compress
        }
        Write-Log ("Request body: {0}" -f $bodyJson) "INFO"
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
            # Parse response body as JSON and return
            if ($webResponse.Content) {
                return $webResponse.Content | ConvertFrom-Json
            }
            return $true
        } catch {
            $status  = $null
            $detail  = $_.Exception.Message
            try { $status = $_.Exception.Response.StatusCode.value__ } catch {}
            try { $detail = $_.ErrorDetails.Message } catch {}
            Write-Log ("API {0} {1} — attempt {2}/{3} failed (HTTP {4}): {5}" -f $Method, $EndpointPath, $attempt, $ApiRetryCount, $status, $detail) "WARN"
            if ($attempt -lt $ApiRetryCount) { Start-Sleep -Milliseconds $ApiRetryDelayMs }
        }
    }
    Write-Log ("ERROR: API call failed after {0} attempts: {1} {2}" -f $ApiRetryCount, $Method, $EndpointPath) "ERROR"
    return $null
}

# =====================================================
# Unit suggestion helper for WMI properties
# Guesses a sensible unit based on property name patterns.
# User is always shown the suggestion and can confirm or override.
# =====================================================
function Get-SuggestedUnit {
    param([string]$PropertyName)
    $n = $PropertyName.ToLower()
    if     ($n -match "latency|roundtrip|rtt|delay|responsetime|duration") { return "ms"         }
    elseif ($n -match "bytespersec|bytespersecon|bytes.*sec")               { return "Bytes/sec"  }
    elseif ($n -match "bytes$")                                             { return "Bytes"      }
    elseif ($n -match "frame|fps|framespersec")                             { return "Frames/sec" }
    elseif ($n -match "percent|usage|utilization|util$|pct$")              { return "%"          }
    elseif ($n -match "mb$|megabyte|availablemb|totalmb")                  { return "MB"         }
    elseif ($n -match "kb$|kilobyte")                                       { return "KB"         }
    elseif ($n -match "mhz|frequency|clock")                               { return "MHz"        }
    elseif ($n -match "watt|power|mw$|milliwatt")                          { return "mW"         }
    elseif ($n -match "count$|length$|queue|sessions?$|connections?$|packets?$") { return "Count" }
    elseif ($n -match "persec$|persecond$|rate$")                          { return "/sec"       }
    elseif ($n -match "temp|celsius")                                       { return "C"          }
    elseif ($n -match "mbps|megabit")                                       { return "Mbps"       }
    else                                                                    { return "%"          }
}
# =====================================================
Write-Console ""
Write-Console "============================================================" Cyan
Write-Console " Login Enterprise — Session Metric Discovery" Cyan
Write-Console " Searches PerfMon and WMI on this machine for performance" Cyan
Write-Console " counters matching your search string, then uploads them" Cyan
Write-Console " as session metric definitions to your LE appliance." Cyan
Write-Console "============================================================" Cyan
Write-Console ""

if ($SearchString -eq "") {
    Write-Console "EXAMPLE SEARCH STRINGS" Yellow
    Write-Console ""
    Write-Console "  Remoting protocol / display protocol counters:" Yellow
    Write-Console "    RemoteFX       Microsoft RemoteFX graphics (RDP / Azure Virtual Desktop)"
    Write-Console "    VMwareBlast    Omnissa (VMware Horizon) Blast Extreme protocol"
    Write-Console "    Blast          Broader Omnissa Blast counter search"
    Write-Console "    PCoIP          Teradici PCoIP protocol (Omnissa/AWS)"
    Write-Console "    Citrix         Citrix ICA / HDX related counters"
    Write-Console "    EUEM           Citrix End User Experience Monitoring (WMI)"
    Write-Console "    HDX            Citrix HDX protocol counters"
    Write-Console ""
    Write-Console "  GPU / graphics:" Yellow
    Write-Console "    GPUEngine      WMI GPU engine utilization (vendor-agnostic)"
    Write-Console "    NVIDIA         NVIDIA PerfMon GPU counters"
    Write-Console "    GPU            Broad GPU counter search across PerfMon and WMI"
    Write-Console ""
    Write-Console "  System / network:" Yellow
    Write-Console "    Network Interface   NIC throughput, packets, errors"
    Write-Console "    TCPv4               TCP connection and throughput"
    Write-Console "    PhysicalDisk        Disk throughput, latency, queue"
    Write-Console "    Memory              Memory usage, paging"
    Write-Console "    Processor           CPU utilization variants"
    Write-Console ""
    Write-Console "  NOTE: The correct counter name for total framerate metrics" Magenta
    Write-Console "  (Omnissa Blast, Citrix) depends on which protocol agent is" Magenta
    Write-Console "  installed on this machine. Try 'RemoteFX', 'VMwareBlast'," Magenta
    Write-Console "  or 'Citrix' to explore what is available." Magenta
    Write-Console ""

    $SearchString = Read-Host "Enter search string"
    if ($SearchString.Trim() -eq "") {
        Write-Log "No search string provided. Exiting." "ERROR"
        exit 1
    }
    $SearchString = $SearchString.Trim()
}

Write-Log ("Search string: '{0}'" -f $SearchString)

# =====================================================
# Step 2 — Search PerfMon
# =====================================================
Write-Log "------------------------------------------------------------"
Write-Log ("Step 2: Searching PerfMon counter sets for '{0}'..." -f $SearchString)

$perfmonResults = @()

try {
    $matchingSets = Get-Counter -ListSet * -ErrorAction SilentlyContinue |
                    Where-Object { $_.CounterSetName -match [regex]::Escape($SearchString) }

    foreach ($set in $matchingSets) {
        # Paths with instances (has an instance in parens)
        foreach ($path in $set.PathsWithInstances) {
            if ($path -match '\\(.+?)\((.+?)\)\\(.+)') {
                $perfmonResults += [PSCustomObject]@{
                    Source      = "PerfMon"
                    CounterSet  = $set.CounterSetName
                    CounterName = $Matches[3]
                    Instance    = $Matches[2]
                    FullPath    = $path
                }
            }
        }
        # Paths without instances
        foreach ($path in ($set.Paths | Where-Object { $_ -notmatch '\(' })) {
            if ($path -match '\\[^\\]+\\(.+)$') {
                $perfmonResults += [PSCustomObject]@{
                    Source      = "PerfMon"
                    CounterSet  = $set.CounterSetName
                    CounterName = $Matches[1]
                    Instance    = ""
                    FullPath    = $path
                }
            }
        }
    }
} catch {
    Write-Log ("PerfMon search warning: {0}" -f $_.Exception.Message) "WARN"
}

Write-Log ("{0} PerfMon counter(s) found" -f $perfmonResults.Count)

# =====================================================
# Step 3 — Search WMI
# =====================================================
Write-Log "------------------------------------------------------------"
Write-Log ("Step 3: Searching WMI classes in {0} for '{1}'..." -f $WmiNamespace, $SearchString)

$wmiResults = @()

try {
    $matchingClasses = Get-WmiObject -Namespace $WmiNamespace -List -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -match [regex]::Escape($SearchString) }

    foreach ($class in $matchingClasses) {
        $props = $class.Properties |
                 Where-Object { $WmiSystemProperties -notcontains $_.Name } |
                 Select-Object -ExpandProperty Name | Sort-Object

        if ($props -and @($props).Count -gt 0) {
            $wmiResults += [PSCustomObject]@{
                Source     = "WMI"
                ClassName  = $class.Name
                Properties = @($props)
                Namespace  = $WmiNamespace
            }
        }
    }
} catch {
    Write-Log ("WMI search warning: {0}" -f $_.Exception.Message) "WARN"
}

Write-Log ("{0} WMI class(es) found" -f $wmiResults.Count)

# =====================================================
# Step 4 — Present numbered list
# =====================================================
if ($perfmonResults.Count -eq 0 -and $wmiResults.Count -eq 0) {
    Write-Console ""
    Write-Console ("No PerfMon counters or WMI classes found matching '{0}'." -f $SearchString) Red
    Write-Console "Try a different or broader search string. See examples above." Yellow
    Write-Log "No results found. Exiting." "INFO"
    exit 0
}

Write-Console ""
Write-Console "============================================================" Cyan
Write-Console (" Search results for: '{0}'" -f $SearchString) Cyan
Write-Console "============================================================" Cyan

$displayItems = @()
$itemNumber   = 1

if ($perfmonResults.Count -gt 0) {
    Write-Console ""
    Write-Console "  PERFMON COUNTERS" Yellow

    $grouped = $perfmonResults | Group-Object CounterSet
    foreach ($group in $grouped) {
        Write-Console ("  [ {0} ]" -f $group.Name) Cyan
        $unique = $group.Group | Sort-Object CounterName | Select-Object CounterSet, CounterName, Instance -Unique
        foreach ($item in $unique) {
            $instDisplay = if ($item.Instance -ne "") { "  (instance: {0})" -f $item.Instance } else { "" }
            Write-Console ("    {0,3}.  {1}{2}" -f $itemNumber, $item.CounterName, $instDisplay)
            $displayItems += [PSCustomObject]@{
                Number      = $itemNumber
                Source      = "PerfMon"
                CounterSet  = $item.CounterSet
                CounterName = $item.CounterName
                Instance    = $item.Instance
            }
            $itemNumber++
        }
    }
}

if ($wmiResults.Count -gt 0) {
    Write-Console ""
    Write-Console "  WMI CLASSES" Yellow
    foreach ($wmiClass in $wmiResults) {
        Write-Console ("  [ {0} ]" -f $wmiClass.ClassName) Cyan
        Write-Console ("    {0,3}.  Add from this class  ({1} available properties — you will choose next)" -f $itemNumber, $wmiClass.Properties.Count)
        $displayItems += [PSCustomObject]@{
            Number     = $itemNumber
            Source     = "WMI"
            ClassName  = $wmiClass.ClassName
            Properties = $wmiClass.Properties
            Namespace  = $wmiClass.Namespace
        }
        $itemNumber++
    }
}

Write-Console ""

# =====================================================
# Step 5 — User selects items
# =====================================================
Write-Console "Enter the number(s) of the counters / classes to add." White
Write-Console "Separate multiple with commas.  Example: 1,3,4" White
Write-Console ""
$selectionInput = (Read-Host "Your selection").Trim()

if ($selectionInput -eq "") {
    Write-Log "No selection entered. Exiting." "INFO"
    exit 0
}

$selectedNumbers = $selectionInput -split ',' |
                   ForEach-Object { $_.Trim() } |
                   Where-Object   { $_ -match '^\d+$' } |
                   ForEach-Object { [int]$_ } |
                   Select-Object -Unique

$selectedItems = $displayItems | Where-Object { $selectedNumbers -contains $_.Number }

if ($selectedItems.Count -eq 0) {
    Write-Log "No valid items matched the selection. Exiting." "WARN"
    exit 1
}

# =====================================================
# Step 6 — For WMI items, choose properties
# =====================================================
$metricsToCreate = @()

foreach ($item in $selectedItems) {

    if ($item.Source -eq "PerfMon") {
        $metricName = "{0} - {1}" -f $item.CounterSet, $item.CounterName
        $suggested  = Get-SuggestedUnit -PropertyName $item.CounterName
        Write-Console ""
        Write-Console ("  Unit for: {0}" -f $metricName) Cyan
        Write-Console "  (press Enter to accept suggestion, or type your own e.g. ms, Bytes/sec, Frames/sec, Count, MB)" Gray
        $unitInput  = (Read-Host ("  unit [{0}]" -f $suggested)).Trim()
        $unit       = if ($unitInput -eq "") { $suggested } else { $unitInput }
        Write-Log ("  PerfMon counter: {0}  unit: {1}" -f $metricName, $unit)

        $metricsToCreate += [PSCustomObject]@{
            Source     = "PerfMon"
            Label      = $metricName
            Definition = @{
                type        = "PerformanceCounter"
                name        = $metricName
                description = "PerfMon counter: {0} \ {1}{2}" -f $item.CounterSet, $item.CounterName, $(if ($item.Instance -ne "") { " ({0})" -f $item.Instance } else { "" })
                tag         = $MetricTag
                measurement = @{
                    counterCategory = $item.CounterSet
                    counterName     = $item.CounterName
                    counterInstance = $item.Instance
                    displayName     = $metricName
                    unit            = $unit
                }
            }
        }
    }

    elseif ($item.Source -eq "WMI") {
        Write-Console ""
        Write-Console ("  Properties in [{0}]:" -f $item.ClassName) Cyan
        Write-Console "      0.  Add ALL properties"
        $pNum = 1
        foreach ($prop in $item.Properties) {
            Write-Console ("    {0,3}.  {1}" -f $pNum, $prop)
            $pNum++
        }
        Write-Console ""
        Write-Console "  Enter property number(s) to add, or 0 for all. Separate with commas." White
        $propInput = (Read-Host ("  Properties for {0}" -f $item.ClassName)).Trim()

        $selectedProps = @()
        if ($propInput -eq "0" -or $propInput -eq "") {
            $selectedProps = $item.Properties
            Write-Log ("All {0} properties selected from {1}" -f $selectedProps.Count, $item.ClassName)
        } else {
            $propNums      = $propInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
            $selectedProps = $propNums | ForEach-Object {
                $idx = $_ - 1
                if ($idx -ge 0 -and $idx -lt $item.Properties.Count) { $item.Properties[$idx] }
            } | Where-Object { $_ -ne $null }
        }

        if ($selectedProps.Count -eq 0) {
            Write-Log ("No valid properties selected for {0} — skipping." -f $item.ClassName) "WARN"
            continue
        }

        # Per-property unit selection with smart suggestion
        Write-Console ""
        Write-Console "  Unit selection for each property:" Cyan
        Write-Console "  (press Enter to accept suggestion, or type your own e.g. ms, Bytes/sec, Frames/sec, Count, MB)" Gray
        Write-Console ""

        $measurements = @()
        foreach ($prop in $selectedProps) {
            $suggested = Get-SuggestedUnit -PropertyName $prop
            $unitInput = (Read-Host ("  {0,-45} unit [{1}]" -f $prop, $suggested)).Trim()
            $unit      = if ($unitInput -eq "") { $suggested } else { $unitInput }

            # summarizeOperation — suggest based on property name
            $suggestedOp = if ($prop -match "Max|Peak|Highest")    { "max" }
                           elseif ($prop -match "Total|Sum|Count") { "sum" }
                           elseif ($prop -match "Rate|Per")        { "none" }
                           else                                    { "avg" }
            Write-Console "    summarize: avg=average over interval, sum=total, max=peak, none=last raw value" Gray
            $opInput = (Read-Host ("  {0,-45} summarize [{1}]" -f $prop, $suggestedOp)).Trim().ToLower()
            $op      = if ($opInput -in @("avg","sum","max","min","none")) { $opInput } else { $suggestedOp }

            Write-Log ("  Property: {0}  unit: {1}  summarize: {2}" -f $prop, $unit, $op)

            # displayName must be unique and <=64 chars
            # Use last segment of class name (after last underscore) as short prefix
            $classShort  = $item.ClassName.Split('_') | Select-Object -Last 1
            if (-not $classShort) { $classShort = $item.ClassName }
            $candidate   = "{0} - {1} (WMI)" -f $classShort, $prop
            $displayName = if ($candidate.Length -le 64) { $candidate } else { $candidate.Substring(0, 64) }

            $measurements += @{
                propertyName       = $prop
                summarizeOperation = $op
                displayName        = $displayName
                unit               = $unit
            }
        }

        $wmiQuery = "SELECT {0} FROM {1}" -f (($selectedProps) -join ", "), $item.ClassName

        $metricsToCreate += [PSCustomObject]@{
            Source     = "WMI"
            Label      = ("{0} ({1} properties)" -f $item.ClassName, $selectedProps.Count)
            Definition = @{
                type          = "WmiQuery"
                name          = $item.ClassName
                description   = "WMI class: {0} | Namespace: {1}" -f $item.ClassName, $item.Namespace
                tag           = $MetricTag
                wmiQuery      = $wmiQuery
                namespace     = $item.Namespace
                instanceField = "Name"
                measurements  = @($measurements)
            }
        }
    }
}

if ($metricsToCreate.Count -eq 0) {
    Write-Log "Nothing to create after property selection. Exiting." "INFO"
    exit 0
}

# =====================================================
# Step 7 — Confirmation
# =====================================================
Write-Console ""
Write-Console "============================================================" Cyan
Write-Console " Confirm — will be created in Login Enterprise:" Cyan
Write-Console "============================================================" Cyan
foreach ($m in $metricsToCreate) {
    Write-Console ("  [{0}]  {1}" -f $m.Source, $m.Label)
    if ($m.Source -eq "WMI") {
        Write-Console ("          {0}" -f $m.Definition.wmiQuery) Gray
    }
}
Write-Console ""
$confirm = (Read-Host "Proceed? (Y/N)").Trim()
if ($confirm -notmatch '^[Yy]') {
    Write-Log "Cancelled by user. No changes made." "INFO"
    exit 0
}

# =====================================================
# Step 8 — Create metric definitions in LE
# =====================================================
Write-Log "------------------------------------------------------------"
Write-Log ("Step 8: Creating {0} session metric definition(s) in Login Enterprise..." -f $metricsToCreate.Count)

$createdKeys   = @()
$createdLabels = @()

foreach ($m in $metricsToCreate) {
    Write-Log ("  Creating [{0}]: {1}" -f $m.Source, $m.Label)
    $result = Invoke-LeApi -Method "POST" -EndpointPath "user-session-metric-definitions" -Body $m.Definition

    if ($result -and $result.id) {
        Write-Log ("    [OK] ID: {0}" -f $result.id)
        $createdKeys   += $result.id
        $createdLabels += $m.Label
    } else {
        Write-Log ("    [ERROR] Failed to create: {0}" -f $m.Label) "ERROR"
    }
}

if ($createdKeys.Count -eq 0) {
    Write-Log "No metric definitions were created successfully. Exiting." "ERROR"
    exit 1
}

Write-Log ("{0}/{1} metric definition(s) created." -f $createdKeys.Count, $metricsToCreate.Count)

# =====================================================
# Step 9 — Group assignment
# =====================================================
Write-Console ""
Write-Console "============================================================" Cyan
Write-Console " Group Assignment" Cyan
Write-Console "============================================================" Cyan
Write-Console ""

$groupAction     = $null
$finalGroupId    = $null
$finalGroupName  = $null

if ($ExistingGroupId -ne "") {
    $groupAction = "existing"
    # Fetch the group name for summary display
    $existingGroup = Invoke-LeApi -Method "GET" -EndpointPath ("session-metric-definition-groups/{0}" -f $ExistingGroupId)
    if ($existingGroup -and $existingGroup.name) { $finalGroupName = $existingGroup.name }
    else { $finalGroupName = $ExistingGroupId }
} elseif ($NewGroupName -ne "") {
    $groupAction = "new"
} else {
    Write-Console "What would you like to do with the created metric(s)?" White
    Write-Console "  1.  Add to an existing session metric group"
    Write-Console "  2.  Create a new group containing these metrics"
    Write-Console "  3.  Skip — leave metrics unassigned (assign later in LE UI)"
    Write-Console ""
    $groupChoice = (Read-Host "Enter 1, 2, or 3").Trim()
    switch ($groupChoice) {
        "1"     { $groupAction = "existing" }
        "2"     { $groupAction = "new"      }
        default { $groupAction = "skip"     }
    }
}

# --- Add to existing group ---
if ($groupAction -eq "existing") {
    if ($ExistingGroupId -eq "") {
        Write-Console ""
        Write-Console "Enter a group ID (UUID) or partial group name to search:" White
        $groupInput = (Read-Host "Group ID or name").Trim()

        if ($groupInput -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
            $ExistingGroupId = $groupInput
            Write-Log ("Using group ID directly: {0}" -f $ExistingGroupId)
        } else {
            Write-Log ("Searching for group matching '{0}'..." -f $groupInput)
            $groupSearch = Invoke-LeApi -Method "GET" -EndpointPath "session-metric-definition-groups" `
                -Query @{ filter = $groupInput; count = 20; offset = 0 }

            if ($groupSearch -and $groupSearch.items -and $groupSearch.items.Count -gt 0) {
                if ($groupSearch.items.Count -eq 1) {
                    $ExistingGroupId = $groupSearch.items[0].key
                    $finalGroupName  = $groupSearch.items[0].name
                    Write-Log ("Found group: '{0}' — ID: {1}" -f $finalGroupName, $ExistingGroupId)
                } else {
                    Write-Console ""
                    Write-Console "Multiple groups found — select one:" White
                    $gNum = 1
                    foreach ($g in $groupSearch.items) {
                        Write-Console ("  {0,3}.  {1}  ({2} members)" -f $gNum, $g.name, $g.memberCount)
                        $gNum++
                    }
                    $gChoice = (Read-Host "Group number").Trim()
                    $gIdx    = ([int]$gChoice) - 1
                    if ($gIdx -ge 0 -and $gIdx -lt $groupSearch.items.Count) {
                        $ExistingGroupId = $groupSearch.items[$gIdx].key
                        $finalGroupName  = $groupSearch.items[$gIdx].name
                        Write-Log ("Selected: '{0}' — ID: {1}" -f $finalGroupName, $ExistingGroupId)
                    } else {
                        Write-Log "Invalid selection — skipping group assignment." "WARN"
                        $groupAction = "skip"
                    }
                }
            } else {
                Write-Log ("No group found matching '{0}' — skipping group assignment." -f $groupInput) "WARN"
                $groupAction = "skip"
            }
        }
    }

    if ($groupAction -eq "existing" -and $ExistingGroupId -ne "") {
        Write-Log ("Adding {0} metric(s) to group {1}..." -f $createdKeys.Count, $ExistingGroupId)
        $addResult = Invoke-LeApi -Method "POST" `
            -EndpointPath ("session-metric-definition-groups/{0}/members" -f $ExistingGroupId) `
            -Body $createdKeys
        if ($null -ne $addResult) {
            Write-Log ("[OK] Metrics added to group.")
        }
        $finalGroupId = $ExistingGroupId
    }
}

# --- Create new group ---
elseif ($groupAction -eq "new") {
    if ($NewGroupName -eq "") {
        $NewGroupName = (Read-Host "New group name").Trim()
    }
    if ($NewGroupDescription -eq "") {
        $NewGroupDescription = "Created by Add-SessionMetrics.ps1 v$ScriptVersion on {0}" -f (Get-Date).ToString("yyyy-MM-dd")
    }

    Write-Log ("Creating new group '{0}'..." -f $NewGroupName)
    $newGroupResult = Invoke-LeApi -Method "POST" -EndpointPath "session-metric-definition-groups" -Body @{
        name           = $NewGroupName
        description    = $NewGroupDescription
        definitionKeys = $createdKeys
    }

    if ($newGroupResult -and $newGroupResult.id) {
        $finalGroupId   = $newGroupResult.id
        $finalGroupName = $NewGroupName
        Write-Log ("[OK] New group created — ID: {0}" -f $finalGroupId)
    } else {
        Write-Log "ERROR: Failed to create new group. Metrics were created but not grouped." "ERROR"
    }
}

# --- Skip ---
else {
    Write-Log "Group assignment skipped."
}

# =====================================================
# Summary
# =====================================================
Write-Console ""
Write-Console "============================================================" Cyan
Write-Console " DONE — Summary" Cyan
Write-Console "============================================================" Cyan
Write-Log ""
Write-Log ("Metric definitions created: {0}" -f $createdKeys.Count)
for ($i = 0; $i -lt $createdLabels.Count; $i++) {
    $key = if ($i -lt $createdKeys.Count) { $createdKeys[$i] } else { "n/a" }
    Write-Log ("  [{0}] {1}  —  key: {2}" -f $metricsToCreate[$i].Source, $createdLabels[$i], $key)
}
Write-Log ""
if ($finalGroupId) {
    Write-Log ("Group: '{0}'  —  ID: {1}" -f $finalGroupName, $finalGroupId)
} else {
    Write-Log "No group assigned. Assign later via LE UI: Configuration > Session Metrics > Groups"
}
Write-Log ""
Write-Log "NEXT STEPS:"
Write-Log "  1. Verify metrics: LE UI > Configuration > Session Metrics > Metrics"
Write-Log "     Review unit, display name — edit if needed (delete and recreate, cannot edit in LE)"
Write-Log "  2. If no group assigned, add to group: Configuration > Session Metrics > Groups"
Write-Log "  3. Assign group to your test: Configuration > Manage Tests > [Test] > Session Metrics"
Write-Log "  4. Run a short no-action test to confirm all metrics are collecting."
Write-Log ("Log saved to: {0}" -f $LogFile)
Write-Console "============================================================" Cyan


# Cert cleanup
if ($ImportServerCert -and -not $KeepCert -and $script:ImportedCertThumbs.Length -gt 0) {
    try { Remove-ImportedCertificates -Thumbprints $script:ImportedCertThumbs }
    catch { Write-Log ("Cert cleanup failed: {0}" -f $_.Exception.Message) "WARN" }
}
