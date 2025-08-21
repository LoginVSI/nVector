<#
nVector_Client_Prepare.ps1
version 1.0.7 -- Robust WallClock parsing: Try Parse() first, TryParseExact loop as fallback (no multi-format overload)
Change log (high level):
  1.0.0 - original combined uploader logic
  1.0.1 - separated server/local drift helper; initial tests
  1.0.2 - introduced RTT / RawUtc / WallClock modes
  1.0.3 - Normalize +0000 -> +00:00 handling
  1.0.4 - Added SanityMaxHours + improved logging
  1.0.5 - WallClock default + ForceLocalOffset option
  1.0.6 - Made WallClock parsing more tolerant (TryParseExact + fallback to Parse)
  1.0.7 - **Flip parsing order**: Try Parse() first (working path), only then try TryParseExact loop (fallback); avoids the TryParseExact multi-format overload errors in PS5.
#>

# -----------------------------
# Config (edit as needed)
# -----------------------------
$ScriptVersion = "1.0.7"
$AdjustmentMode = "WallClock"   # "WallClock" (default), "RTT", or "RawUtc"
$ForceLocalOffset = $null       # e.g. "-07:00" to force Pacific behavior for wall-clock mode, otherwise $null
$SanityMaxHours = 168           # guard: abort if absolute adjustment > this (0 disables)

$NvectorAgentCheckIntervalMs = 5000  
$CsvFilePath          = "C:\temp\nvidia\latency_metrics.csv"
$NvectorScreenshotDir = "C:\temp\nvidia\SSIM_screenshots"
$NvectorLogFile       = "C:\temp\nvidia\agent.log"

$Timestamp            = (Get-Date).ToString('yyyyMMddTHHmmss')
$TranscriptFile       = "C:\temp\nvidia\${Timestamp}nVector_Agent_Client.log"
$ErrorActionPreference = 'Continue'
$VerbosePreference     = 'Continue'
$DebugPreference       = 'Continue'
Start-Transcript -Path $TranscriptFile -Append

$Arguments = @(
    "-r","client",
    "-m", $CsvFilePath,
    "-p", "$NvectorAgentCheckIntervalMs",
    "-s", $NvectorScreenshotDir,
    "-l", $NvectorLogFile
)

$NvectorAgentExePath = "" # Place full path to nvector-agent.exe here

$PollingInterval     = 10
$MaxLatencyThreshold = 1500

$LauncherProcessName = "LoginEnterprise.Launcher.UI"
$LauncherExePath     = "C:\Program Files\Login VSI\Login Enterprise Launcher\LoginEnterprise.Launcher.UI.exe"

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

# -----------------------------
# Helpers
# -----------------------------
function Get-EffectiveLocalDto {
    if ($null -ne $ForceLocalOffset -and $ForceLocalOffset -ne "") {
        try {
            $ts = [TimeSpan]::Parse($ForceLocalOffset)
        } catch {
            Write-Error ("Invalid ForceLocalOffset '{0}' -- must be like '-07:00' or '+01:00'." -f $ForceLocalOffset)
            throw
        }
        $nowUtc = [DateTimeOffset]::UtcNow
        return $nowUtc.ToOffset($ts)
    } else {
        return [DateTimeOffset]::Now
    }
}

function Format-TimeSpanCompact {
    param([TimeSpan]$ts)
    $h = [int][math]::Floor([math]::Abs($ts.TotalHours))
    $m = [int]([math]::Abs($ts.Minutes))
    $s = [int]([math]::Abs($ts.Seconds))
    $parts = @()
    if ($h -ne 0) { if ($h -eq 1) { $hs = "" } else { $hs = "s" }; $parts += ("{0} hour{1}" -f $h, $hs) }
    if ($m -ne 0) { if ($m -eq 1) { $ms = "" } else { $ms = "s" }; $parts += ("{0} minute{1}" -f $m, $ms) }
    if ($s -ne 0) { if ($s -eq 1) { $ss = "" } else { $ss = "s" }; $parts += ("{0} second{1}" -f $s, $ss) }
    if ($parts.Count -eq 0) { return "0 seconds" }
    return ($parts -join ", ")
}

# -----------------------------
# Compute adjustments (candidates)
# -----------------------------
function Compute-AdjustmentCandidates {
    param([string]$BaseUrl, [string]$Endpoint, [string]$Token, [int]$TimeoutSec = 5)

    $uri = $BaseUrl.TrimEnd('/') + '/' + $Endpoint.TrimStart('/')
    $hdr = @{ Authorization = "Bearer $Token" }

    try {
        Write-Host (("Querying server (initial) at {0} ..." -f $uri))
        try { $r1 = Invoke-WebRequest -Uri $uri -Method Head -Headers $hdr -UseBasicParsing -TimeoutSec $TimeoutSec }
        catch { Write-Host "HEAD initial failed, falling back to GET..."; $r1 = Invoke-WebRequest -Uri $uri -Method Get -Headers $hdr -UseBasicParsing -TimeoutSec $TimeoutSec }

        $date1 = $r1.Headers['Date']
        if (-not $date1) { throw "No Date header returned (initial)" }

        Write-Host "Server raw Date header (initial):"
        Write-Host $date1
        $serverDto1 = [DateTimeOffset]::Parse($date1)
        $serverUtc1 = $serverDto1.UtcDateTime
        Write-Host (("Parsed server initial UTC: {0}" -f $serverUtc1.ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'")))

        # effective local
        $localDto = Get-EffectiveLocalDto
        $localUtc = $localDto.UtcDateTime
        Write-Host "`nEffective local snapshot (wallclock):"
        $localDto | Format-List
        Write-Host (("Local formatted (with offset): {0}" -f $localDto.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz")))

        # Candidate: RawUtc
        $adjustRawUtc = $serverUtc1 - $localUtc
        Write-Host (("Candidate RawUtc adjustment (serverUtc_initial - localUtc): {0} ({1} ms)" -f $adjustRawUtc.ToString(), [math]::Round($adjustRawUtc.TotalMilliseconds)))

        # Candidate: WallClock
        $serverAsLocal = $serverDto1.ToOffset($localDto.Offset)
        $adjustWallClock = $serverAsLocal - $localDto
        Write-Host (("Candidate WallClock adjustment (server as local - local wallclock): {0} ({1} ms)" -f $adjustWallClock.ToString(), [math]::Round($adjustWallClock.TotalMilliseconds)))

        # RTT candidate
        Write-Host "`nMeasuring RTT for midpoint estimate..."
        $localBefore = [DateTimeOffset]::UtcNow
        try { $r2 = Invoke-WebRequest -Uri $uri -Method Head -Headers $hdr -UseBasicParsing -TimeoutSec $TimeoutSec }
        catch { Write-Host "HEAD drift request failed, falling back to GET..."; $r2 = Invoke-WebRequest -Uri $uri -Method Get -Headers $hdr -UseBasicParsing -TimeoutSec $TimeoutSec }
        $localAfter = [DateTimeOffset]::UtcNow

        $date2 = $r2.Headers['Date']
        if (-not $date2) { throw "No Date header returned (drift request)" }

        $serverDto2 = [DateTimeOffset]::Parse($date2)
        $serverUtc2 = $serverDto2.UtcDateTime
        Write-Host (("Parsed server drift-request UTC: {0}" -f $serverUtc2.ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'")))

        $roundtrip = $localAfter - $localBefore
        $halfRt = [TimeSpan]::FromTicks([Math]::Floor($roundtrip.Ticks / 2))
        $estimatedLocalUtc = ($localBefore + $halfRt).UtcDateTime
        Write-Host (("RTT roundtrip: {0} ms" -f [math]::Round($roundtrip.TotalMilliseconds)))
        Write-Host (("Estimated local UTC (midpoint): {0:yyyy-MM-ddTHH:mm:ss.fff'Z'}" -f $estimatedLocalUtc))

        $adjustRtt = $serverUtc2 - $estimatedLocalUtc
        Write-Host (("Candidate RTT adjustment (serverUtc_at_midpoint - estimatedLocalUtc): {0} ({1} ms)" -f $adjustRtt.ToString(), [math]::Round($adjustRtt.TotalMilliseconds)))

        return [PSCustomObject]@{
            ServerDtoInitial   = $serverDto1
            ServerUtcInitial   = $serverUtc1
            ServerDtoDriftReq  = $serverDto2
            ServerUtcDriftReq  = $serverUtc2
            LocalDto           = $localDto
            LocalUtc           = $localUtc
            AdjustRawUtc       = $adjustRawUtc
            AdjustWallClock    = $adjustWallClock
            AdjustRtt          = $adjustRtt
            Roundtrip          = $roundtrip
            EstimatedLocalUtc  = $estimatedLocalUtc
        }
    } catch {
        Write-Error ("Compute-AdjustmentCandidates failed: {0}" -f $_.Exception.Message)
        throw
    }
}

# Compute candidates and select by mode
Write-Host (("Selecting adjustment mode: {0}" -f $AdjustmentMode))
$adjInfo = Compute-AdjustmentCandidates -BaseUrl $BaseUrl -Endpoint "v8-preview/system/version" -Token $ConfigurationAccessToken
if (-not $adjInfo) { Write-Error "Adjustment computation failed"; Stop-Transcript; exit 1 }

switch ($AdjustmentMode.ToUpperInvariant()) {
    "WALLCLOCK" {
        $script:AdjustToServer = $adjInfo.AdjustWallClock
        $chosenDesc = "Wall-clock (server-as-local minus local wallclock)"
    }
    "RTT" {
        $script:AdjustToServer = $adjInfo.AdjustRtt
        $chosenDesc = "RTT-compensated (server at midpoint minus estimated local midpoint)"
    }
    "RAWUTC" {
        $script:AdjustToServer = $adjInfo.AdjustRawUtc
        $chosenDesc = "Raw UTC (server initial UTC minus local UTC snapshot)"
    }
    Default {
        Write-Host (("Unknown mode '{0}', defaulting to WallClock." -f $AdjustmentMode))
        $script:AdjustToServer = $adjInfo.AdjustWallClock
        $chosenDesc = "Wall-clock (server-as-local minus local wallclock)"
    }
}

$adjMs = [math]::Round($script:AdjustToServer.TotalMilliseconds)
Write-Host (("Selected adjustment: {0} => {1} ({2} ms)" -f $AdjustmentMode, $script:AdjustToServer.ToString(), $adjMs))

if (($SanityMaxHours -gt 0) -and ([math]::Abs($script:AdjustToServer.TotalHours) -gt $SanityMaxHours)) {
    Write-Error (("Adjustment exceeds sanity cap of {0} hours. Aborting." -f $SanityMaxHours))
    Stop-Transcript; exit 1
}

# For compatibility
$script:TimeOffsetSpan = $script:AdjustToServer

# -----------------------------
# Adjust-TimeOffset (WallClock updated parsing order)
# -----------------------------
function Adjust-TimeOffset {
    param([string]$RawTimestamp)

    if (-not $script:AdjustToServer) {
        Write-Error "No `\$script:AdjustToServer available. Compute adjustment first."
        return $null
    }

    $raw = $RawTimestamp.Trim()
    Write-Host (("Raw timestamp input: {0}" -f $raw))

    if ($AdjustmentMode.ToUpperInvariant() -eq "WALLCLOCK") {

        # Strip trailing offsets like Z, +0000, -0100, +00:00, -01:00
        $stripped = $raw -replace '([Zz]|[+-]\d{2}:?\d{2})$',''
        $stripped = $stripped.Trim()
        if ($stripped -ne $raw) {
            Write-Host (("Stripped offset from input for wallclock handling: '{0}' -> '{1}'" -f $raw, $stripped))
        }

        # Accept formats (T or space) with or without milliseconds
        $formats = @(
            "yyyy-MM-ddTHH:mm:ss.fff",
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-dd HH:mm:ss.fff",
            "yyyy-MM-dd HH:mm:ss"
        )
        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        $styles = [System.Globalization.DateTimeStyles]::AssumeLocal
        $dtLocal = $null

        # NEW: Try flexible Parse() first (this was the reliable working path)
        try {
            $dtLocal = [DateTime]::Parse($stripped, $culture, $styles)
            Write-Host ("Parsed local wallclock via Parse(): {0} (Kind: {1})" -f $dtLocal.ToString("yyyy-MM-ddTHH:mm:ss.fff"), $dtLocal.Kind)
        } catch {
            Write-Host "Flexible Parse() failed — trying exact formats (TryParseExact loop)..."
            $ok = $false
            foreach ($fmt in $formats) {
                $ref = New-Object System.Object
                $refDt = [ref]$ref
                # Use single-format TryParseExact overload per iteration (no multi-format overload)
                $success = [DateTime]::TryParseExact($stripped, $fmt, $culture, $styles, [ref]$refDt)
                if ($success) {
                    $dtLocal = $refDt.Value
                    $ok = $true
                    Write-Host ("Parsed local wallclock via TryParseExact (format {0}): {1}" -f $fmt, $dtLocal.ToString("yyyy-MM-ddTHH:mm:ss.fff"))
                    break
                }
            }
            if (-not $ok) {
                Write-Warning (("Failed to parse as local wallclock with Parse() and TryParseExact fallbacks: '{0}'" -f $stripped))
                return $null
            }
        }

        # Build DateTimeOffset with effective local offset
        $effectiveLocalDto = Get-EffectiveLocalDto
        $localOffset = $effectiveLocalDto.Offset

        try {
            $localDtoCandidate = New-Object System.DateTimeOffset ($dtLocal, $localOffset)
        } catch {
            # fallback construction: assemble using constructor params (safe for older PS)
            $localDtoCandidate = New-Object System.DateTimeOffset -ArgumentList ($dtLocal.Year, $dtLocal.Month, $dtLocal.Day, $dtLocal.Hour, $dtLocal.Minute, $dtLocal.Second, $dtLocal.Millisecond, $localOffset)
        }

        Write-Host (("Local DTO candidate (effective offset applied): {0}" -f $localDtoCandidate.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz")))

        # Apply wall-clock AdjustToServer (serverAsLocal - localDto)
        $adjustedLocal = $localDtoCandidate.Add($script:AdjustToServer)

        # Convert to UTC for API
        $adjustedUtc = $adjustedLocal.UtcDateTime
        $finalStr = $adjustedUtc.ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'")

        Write-Host (("Adjust-Timestamp (WallClock) → Input: {0} | AdjustToServer(ms): {1} | Adjusted(server-aligned UTC): {2}" -f $RawTimestamp, [math]::Round($script:AdjustToServer.TotalMilliseconds), $finalStr))

        return $finalStr
    }

    # Non-WallClock (RTT/RawUtc): preserve offsets if present and parse as DateTimeOffset
    if ($raw -match '([+-]\d{4})$') {
        $normalized = $raw -replace '([+-]\d{2})(\d{2})$','$1:$2'
        Write-Host (("Normalized offset form: {0}" -f $normalized))
        $raw = $normalized
    }

    try {
        $dto = [DateTimeOffset]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture)
        $dtUtc = $dto.UtcDateTime
        Write-Host (("Parsed as DateTimeOffset -> UTC: {0}" -f $dtUtc.ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'")))
    } catch {
        Write-Warning (("Invalid timestamp parse for non-wallclock mode: '{0}' : {1}" -f $raw, $_.Exception.Message))
        return $null
    }

    # Apply selected adjust
    $adjustedUtc = $dtUtc.Add($script:AdjustToServer)
    $finalStr = $adjustedUtc.ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'")
    Write-Host (("Adjust-Timestamp → Input: {0} | AdjustToServer(ms): {1} | Adjusted(server-aligned): {2}" -f $RawTimestamp, [math]::Round($script:AdjustToServer.TotalMilliseconds), $finalStr))
    return $finalStr
}

# -----------------------------
# Agent helpers (unchanged)
# -----------------------------
function Terminate-nvector-agent {
    $p = Get-Process -Name "nvector-agent" -ErrorAction SilentlyContinue
    if ($p) { Write-Host "Killing previous nvector-agent"; $p | Stop-Process -Force }
    else  { Write-Host "nvector-agent not running" }
}
function Start-nvector-agent {
    if (-not (Test-Path $NvectorAgentExePath)) {
        Write-Error "nvector-agent.exe not found at $NvectorAgentExePath"; Stop-Transcript; exit 1
    }
    Start-Process -FilePath $NvectorAgentExePath -ArgumentList $Arguments -NoNewWindow -ErrorAction Stop
    Write-Host "nvector-agent started"
}
function Upload-DataToApi {
    param([array]$Metrics)
    $json = $Metrics | ConvertTo-Json -Depth 10 -Compress
    if (-not $json.TrimStart().StartsWith("[")) { $json = "[$json]" }
    Write-Host (("POST payload: {0}" -f $json))
    $hdr = @{
        Authorization = "Bearer $ConfigurationAccessToken"
        "Content-Type" = "application/json"
    }
    try {
        Invoke-RestMethod -Uri ($BaseUrl.TrimEnd('/') + '/' + $ApiEndpoint.TrimStart('/')) `
                          -Method Post -Headers $hdr -Body $json | Out-Null
        Write-Host "Upload succeeded"
    } catch {
        Write-Host (("Upload error: {0}" -f $_))
    }
}

# -----------------------------
# Main Execution (unchanged)
# -----------------------------
Write-Host "Starting nVector metrics uploader - version $ScriptVersion"
Terminate-nvector-agent
Start-nvector-agent

if (-not (Get-Process -Name $LauncherProcessName -ErrorAction SilentlyContinue)) {
    Start-Process -FilePath $LauncherExePath -NoNewWindow -ErrorAction SilentlyContinue
    Write-Host "Launcher started"
} else {
    Write-Host "Launcher already running"
}

# Wait for CSV and check header
$exists = $false
for ($i = 0; $i -lt 5; $i++) {
    if (Test-Path $CsvFilePath) { $exists = $true; break }
    Start-Sleep -Seconds 1
}
if (-not $exists) { Write-Host "CSV not found after wait" }
else {
    $headerLine = Get-Content $CsvFilePath -First 1
    if (-not $headerLine) { Write-Host "CSV exists but is empty" }
    elseif ($headerLine.Trim() -ne "timestamp,latency_ms") { Write-Host (("Expected header 'timestamp,latency_ms' but found: '{0}'" -f $headerLine)) }
    else { Write-Host "CSV ready" }
}

if (Test-Path $CsvFilePath) { $LastLine = (Get-Content $CsvFilePath).Count - 1 } else { $LastLine = 0 }

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
                    Write-Host (("Bad CSV line: {0}" -f $line))
                    continue
                }
                $tsRaw = $parts[0].Trim()
                $lat   = $parts[1].Trim()

                [double]$val = 0.0
                if ([double]::TryParse($lat, [ref]$val) -and $val -lt $MaxLatencyThreshold) {
                    $ts = Adjust-TimeOffset -RawTimestamp $tsRaw
                    if (-not $ts) { Write-Host (("Skipping line due to timestamp parse/adjust error: {0}" -f $line)); continue }
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
                    Write-Host (("Excluded or invalid latency: '{0}'" -f $lat))
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
