<#
nVector_SSIM_Daemon.ps1
version 1.0.2

INTERNAL SCRIPT - do not invoke directly.
Spawned once by nVector_Client_Prepare.ps1 at startup when Mode includes SSIM.
Runs for the lifetime of the main script and is terminated when the main script exits.

Responsibilities:
  - Poll client-side and desktop-side batch folders continuously
  - When a new batch appears on both sides, copy to ssim_queue (timestamped subfolder)
  - Run ssim-tool against the queue copy, parse best SSIM score, upload to LE API
  - Archive heatmaps to HeatmapArchiveFolder after each batch
  - Track processed batches by queue folder presence (copy = processed; no copy = new)
  - On startup: optionally wipe or archive the existing ssim_queue folder (configurable)
  - Runs until terminated by the parent process or Ctrl+C

Session boundary awareness:
  The daemon does NOT track the target window or agent lifecycle at all.
  It simply watches for batch folders to appear. Because each batch copy goes to a
  uniquely timestamped subfolder, batch folder name reuse across sessions is handled
  naturally: a new batch1 with fresh files triggers a fresh copy with a new timestamp.
  The daemon uses the earliest screenshot file's LastWriteTime as the LE upload timestamp.

Queue folder management:
  On startup, existing ssim_queue contents are handled per $QueueStartupBehavior:
    "wipe"    - delete all existing queue subfolders (default; clean slate)
    "archive" - move existing queue subfolders to ssim_queue_archive\<ts>\ before starting
    "keep"    - leave existing queue subfolders; daemon will skip already-queued batches
#>
param(
    # --- Paths ---
    [string]$NvidiaRootPath,
    [string]$SsimToolExePath,
    [string]$TargetHost,
    [string]$HeatmapFolder           = "",
    [string]$HeatmapArchiveFolder    = "",
    [string]$SsimQueueFolder         = "",

    # --- Queue startup behavior ---
    [string]$QueueStartupBehavior    = "wipe",   # "wipe" | "archive" | "keep"

    # --- API ---
    [string]$BaseUrl,
    [string]$ConfigurationAccessToken,
    [string]$SsimEnvironmentId,
    [string]$ApiEndpointSsim         = "publicApi/v8-preview/platform-metrics",
    [int]$ApiRetryCount              = 3,
    [int]$ApiRetryDelayMs            = 2000,

    # --- Metric metadata ---
    [string]$SsimMetricId            = "nVectorSsimMetricId",
    [string]$SsimDisplayName         = "SSIM Score",
    [string]$SsimUnit                = "SSIM",
    [string]$SsimGroup               = "nVector",
    [string]$SsimComponentType       = "vm",
    [string]$Instance                = $env:COMPUTERNAME,

    # --- Thresholds ---
    [double]$MinSsimThreshold        = 0.0,

    # --- Timing ---
    [int]$SsimDaemonPollIntervalSeconds = 5,     # how often to scan for new batch folders
    [int]$SsimBatchSettleSeconds        = 3,     # wait after batch folder appears before copying
    [int]$SsimToolTimeoutSeconds        = 300,
    [int]$SsimPollIntervalSeconds       = 5,     # between ssim-tool output CSV polls

    # --- Logging ---
    [string]$DaemonLogFile           = ""
)

# =============================================================================
# LOGGING
# =============================================================================
$ErrorActionPreference = 'Continue'

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts   = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    $line = "[$ts][$Level][SSIM-Daemon] $Message"
    Microsoft.PowerShell.Utility\Write-Host $line
    if ($DaemonLogFile -ne "") {
        try { Add-Content -Path $DaemonLogFile -Value $line -ErrorAction SilentlyContinue } catch {}
    }
}

function Ensure-Directory {
    param([string]$Path)
    if ($Path -ne "" -and -not (Test-Path $Path)) {
        try { New-Item -ItemType Directory -Path $Path -Force | Out-Null; Write-Log ("Created directory: {0}" -f $Path) }
        catch { Write-Log ("Could not create directory {0}: {1}" -f $Path, $_.Exception.Message) "WARN" }
    }
}

# =============================================================================
# STARTUP
# =============================================================================
Write-Log "nVector SSIM Daemon starting..."
Write-Log ("NvidiaRootPath:      {0}" -f $NvidiaRootPath)
Write-Log ("TargetHost:          {0}" -f $TargetHost)
Write-Log ("SsimQueueFolder:     {0}" -f $SsimQueueFolder)
Write-Log ("HeatmapFolder:       {0}" -f $(if ($HeatmapFolder -ne "") { $HeatmapFolder } else { "(disabled)" }))
Write-Log ("HeatmapArchive:      {0}" -f $(if ($HeatmapArchiveFolder -ne "") { $HeatmapArchiveFolder } else { "(disabled)" }))
Write-Log ("QueueStartupBehavior:{0}" -f $QueueStartupBehavior)
Write-Log ("PollInterval:        {0}s" -f $SsimDaemonPollIntervalSeconds)
Write-Log ("BatchSettleSeconds:  {0}s" -f $SsimBatchSettleSeconds)
Write-Log ("SsimToolTimeout:     {0}s" -f $SsimToolTimeoutSeconds)

if ($SsimQueueFolder -eq "") { $SsimQueueFolder = Join-Path $NvidiaRootPath "ssim_queue" }

Ensure-Directory $SsimQueueFolder
if ($HeatmapArchiveFolder -ne "") { Ensure-Directory $HeatmapArchiveFolder }

# --- Handle existing queue folder contents on startup ---
$existingQueueItems = @(Get-ChildItem -Path $SsimQueueFolder -Directory -ErrorAction SilentlyContinue)
if ($existingQueueItems.Count -gt 0) {
    switch ($QueueStartupBehavior.ToLower()) {
        "wipe" {
            Write-Log ("Wiping {0} existing queue subfolder(s) (QueueStartupBehavior=wipe)..." -f $existingQueueItems.Count)
            foreach ($item in $existingQueueItems) {
                try { Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue }
                catch { Write-Log ("Could not remove {0}: {1}" -f $item.FullName, $_.Exception.Message) "WARN" }
            }
            Write-Log "Queue wiped."
        }
        "archive" {
            $archiveRoot = Join-Path $NvidiaRootPath "ssim_queue_archive"
            $archiveTs   = (Get-Date).ToString('yyyyMMddTHHmmss')
            $archiveDst  = Join-Path $archiveRoot $archiveTs
            Ensure-Directory $archiveDst
            Write-Log ("Archiving {0} existing queue subfolder(s) to {1}..." -f $existingQueueItems.Count, $archiveDst)
            foreach ($item in $existingQueueItems) {
                try { Move-Item -Path $item.FullName -Destination $archiveDst -Force -ErrorAction SilentlyContinue }
                catch { Write-Log ("Could not move {0}: {1}" -f $item.FullName, $_.Exception.Message) "WARN" }
            }
            Write-Log "Queue archived."
        }
        "keep" {
            Write-Log ("Keeping {0} existing queue subfolder(s) (QueueStartupBehavior=keep). Will skip already-queued batches." -f $existingQueueItems.Count)
        }
        default {
            Write-Log ("Unknown QueueStartupBehavior '{0}' - defaulting to keep." -f $QueueStartupBehavior) "WARN"
        }
    }
}

# =============================================================================
# QUEUE STATE
# Key: unique queue key derived from batch folder name + content timestamp
# Value: path to the queue subfolder that was created for this batch
# Populated by scanning ssim_queue for existing _client subfolders on startup
# (relevant when QueueStartupBehavior = "keep")
# =============================================================================
$script:ProcessedQueueKeys = @{}

# Re-register any existing queue entries so they aren't reprocessed
foreach ($dir in @(Get-ChildItem -Path $SsimQueueFolder -Directory -ErrorAction SilentlyContinue)) {
    if ($dir.Name -match '_client$') {
        $key = $dir.Name -replace '_client$',''
        $script:ProcessedQueueKeys[$key] = $dir.FullName
        Write-Log ("Registered existing queue entry: {0}" -f $key)
    }
}

Write-Log "SSIM Daemon ready. Entering polling loop..."

# =============================================================================
# HELPERS
# =============================================================================
function Get-BatchQueueKey {
    # Derive a stable key from the batch folder name and the LastWriteTime of
    # its newest file. Seconds precision is sufficient — files written in the same
    # second are part of the same batch. Milliseconds dropped intentionally to
    # avoid key drift between daemon poll iterations.
    # If the folder is repopulated (agent restart / new session), the newest file
    # timestamp changes and we get a new key, so it's treated as a fresh batch.
    param([string]$BatchPath)
    if (-not (Test-Path $BatchPath)) { return $null }
    $newest = Get-ChildItem -Path $BatchPath -File -Recurse -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($newest) {
        $ts = $newest.LastWriteTime.ToString('yyyyMMddTHHmmss')
    } else {
        $folderItem = Get-Item $BatchPath -ErrorAction SilentlyContinue
        if (-not $folderItem) { return $null }
        $ts = $folderItem.LastWriteTime.ToString('yyyyMMddTHHmmss')
    }
    $folderName = Split-Path $BatchPath -Leaf
    return "${folderName}_${ts}"
}

function Copy-BatchToQueue {
    param([string]$BatchName, [string]$QueueKey)

    $clientSrc  = Join-Path $NvidiaRootPath $BatchName
    $desktopSrc = "\\$TargetHost\C`$\ProgramData\NVIDIA Corporation\nVector\$BatchName"
    $clientDst  = Join-Path $SsimQueueFolder "${QueueKey}_client"
    $desktopDst = Join-Path $SsimQueueFolder "${QueueKey}_desktop"

    Write-Log ("Copying {0} to queue [{1}]..." -f $BatchName, $QueueKey)

    if (-not (Test-Path $clientSrc)) {
        Write-Log ("Client folder not found: {0}" -f $clientSrc) "WARN"; return $null
    }
    if (-not (Test-Path $desktopSrc)) {
        Write-Log ("Desktop UNC not reachable: {0}" -f $desktopSrc) "WARN"; return $null
    }

    try {
        Copy-Item -Path $clientSrc  -Destination $clientDst  -Recurse -Force -ErrorAction Stop
        Copy-Item -Path $desktopSrc -Destination $desktopDst -Recurse -Force -ErrorAction Stop
    } catch {
        Write-Log ("Copy failed for {0}: {1}" -f $BatchName, $_.Exception.Message) "WARN"
        return $null
    }

    # Capture timestamp from earliest screenshot in client copy for LE upload x-axis
    $firstImg = Get-ChildItem -Path $clientDst -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime | Select-Object -First 1
    $captureTs = if ($firstImg) {
        $firstImg.LastWriteTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'")
    } else {
        (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'")
    }

    Write-Log ("Queued {0}. CaptureTs: {1}" -f $BatchName, $captureTs)
    return @{
        BatchName    = $BatchName
        QueueKey     = $QueueKey
        ClientQueue  = $clientDst
        DesktopQueue = $desktopDst
        CaptureTs    = $captureTs
    }
}

function Invoke-SsimToolOnEntry {
    param([hashtable]$Entry)

    $batchName    = $Entry.BatchName
    $queueKey     = $Entry.QueueKey
    $clientQueue  = $Entry.ClientQueue
    $desktopQueue = $Entry.DesktopQueue
    $outputCsv    = Join-Path $NvidiaRootPath "${queueKey}_ssim.csv"
    $outputLog    = Join-Path $NvidiaRootPath "${queueKey}_ssim.log"

    Write-Log ("--- SSIM: {0} [{1}] ---" -f $batchName, $queueKey)

    if (-not (Test-Path $SsimToolExePath)) {
        Write-Log ("ssim-tool.exe not found: {0}" -f $SsimToolExePath) "ERROR"; return $null
    }

    $ssimArgs = @(
        "--client_folder",  "`"$clientQueue`"",
        "--desktop_folder", "`"$desktopQueue`"",
        "--output_file",    "`"$outputCsv`"",
        "--log",            "INFO",
        "--log_file",       "`"$outputLog`""
    )

    $batchHeatmapDir = $null
    if ($HeatmapFolder -ne "") {
        # Use queueKey in heatmap subfolder name so reruns don't collide
        $batchHeatmapDir = Join-Path $HeatmapFolder "${queueKey}"
        Ensure-Directory $batchHeatmapDir
        $ssimArgs += "--heatmap_folder"
        $ssimArgs += "`"$batchHeatmapDir`""
        Write-Log ("  heatmap: {0}" -f $batchHeatmapDir)
    }

    Write-Log ("Launching ssim-tool (timeout: {0}s)..." -f $SsimToolTimeoutSeconds)
    try {
        $ssimProc = Start-Process -FilePath $SsimToolExePath -ArgumentList $ssimArgs -NoNewWindow -PassThru -ErrorAction Stop
    } catch {
        Write-Log ("Failed to launch ssim-tool: {0}" -f $_.Exception.Message) "ERROR"; return $null
    }

    $elapsed = 0; $bestRow = $null
    while ($elapsed -lt $SsimToolTimeoutSeconds) {
        Start-Sleep -Seconds $SsimPollIntervalSeconds
        $elapsed += $SsimPollIntervalSeconds
        $procExited = $false
        try { $procExited = $ssimProc.HasExited } catch {}
        if ($procExited) {
            try { Write-Log ("ssim-tool exited (code {0}) after {1}s" -f $ssimProc.ExitCode, $elapsed) } catch { Write-Log ("ssim-tool exited after {0}s" -f $elapsed) }
            break
        }
        if (Test-Path $outputCsv) {
            try {
                $rows    = Import-Csv -Path $outputCsv -ErrorAction SilentlyContinue
                $bestRow = $rows | Where-Object { $_.ID -eq "best" }
                if ($bestRow) { Write-Log ("'best' row found after {0}s" -f $elapsed); break }
            } catch {}
        }
        Write-Log ("Waiting for ssim-tool... {0}s elapsed" -f $elapsed)
    }
    $procExited = $false
    try { $procExited = $ssimProc.HasExited } catch {}
    if (-not $procExited) {
        Write-Log ("ssim-tool timeout - killing" -f $elapsed) "WARN"
        try { $ssimProc | Stop-Process -Force } catch {}
    }
    if (-not $bestRow -and (Test-Path $outputCsv)) {
        try {
            $rows    = Import-Csv -Path $outputCsv -ErrorAction SilentlyContinue
            $bestRow = $rows | Where-Object { $_.ID -eq "best" }
        } catch {}
    }
    if (-not $bestRow) {
        Write-Log ("No 'best' row in SSIM output for {0}" -f $batchName) "WARN"; return $null
    }

    [double]$ssimVal = 0.0
    if (-not [double]::TryParse($bestRow.SSIM, [ref]$ssimVal) -or $ssimVal -le $MinSsimThreshold) {
        Write-Log ("SSIM value '{0}' invalid or below minimum {1}" -f $bestRow.SSIM, $MinSsimThreshold) "WARN"; return $null
    }
    Write-Log ("SSIM: {0} ({1:P2})" -f $ssimVal, $ssimVal)

    # Archive heatmaps
    if ($null -ne $batchHeatmapDir -and $HeatmapArchiveFolder -ne "" -and (Test-Path $batchHeatmapDir)) {
        $archiveDst = Join-Path $HeatmapArchiveFolder $queueKey
        try {
            Copy-Item -Path $batchHeatmapDir -Destination $archiveDst -Recurse -Force
            Write-Log ("Heatmaps archived: {0}" -f $archiveDst)
        } catch {
            Write-Log ("Heatmap archive failed: {0}" -f $_.Exception.Message) "WARN"
        }
    }

    return @{ SsimVal = $ssimVal; CaptureTs = $Entry.CaptureTs; BatchName = $batchName }
}

function Upload-SsimResult {
    param([hashtable]$Result)
    $metric = [PSCustomObject]@{
        metricId       = $SsimMetricId
        environmentKey = $SsimEnvironmentId
        timestamp      = $Result.CaptureTs
        displayName    = $SsimDisplayName
        unit           = $SsimUnit
        instance       = $Instance
        value          = $Result.SsimVal
        group          = $SsimGroup
        componentType  = $SsimComponentType
    }
    # PS5 ConvertTo-Json collapses a single-element array to an object.
    # The LE API requires a JSON array so we force the brackets explicitly.
    $jsonInner = $metric | ConvertTo-Json -Depth 10 -Compress
    $json = "[$jsonInner]"
    $hdr  = @{ Authorization = "Bearer $ConfigurationAccessToken"; "Content-Type" = "application/json" }
    $url  = $BaseUrl.TrimEnd('/') + '/' + $ApiEndpointSsim.TrimStart('/')
    Write-Log ("Uploading SSIM {0} ({1:P2}) for {2} [ts: {3}]" -f $Result.SsimVal, $Result.SsimVal, $Result.BatchName, $Result.CaptureTs)
    for ($attempt = 1; $attempt -le $ApiRetryCount; $attempt++) {
        try {
            Invoke-RestMethod -Uri $url -Method Post -Headers $hdr -Body $json | Out-Null
            Write-Log ("Upload succeeded (attempt {0})" -f $attempt)
            return $true
        } catch {
            Write-Log ("Upload attempt {0}/{1} failed: {2}" -f $attempt, $ApiRetryCount, $_) "WARN"
            if ($attempt -lt $ApiRetryCount) { Start-Sleep -Milliseconds $ApiRetryDelayMs }
        }
    }
    Write-Log ("Upload failed after {0} attempts." -f $ApiRetryCount) "ERROR"
    return $false
}

# =============================================================================
# MAIN DAEMON LOOP
# Polls for new batch folders on client side. When a new batch appears and its
# corresponding desktop folder is also reachable, copies both to queue and
# processes synchronously (this is a single-threaded daemon; ssim-tool runs
# inline here but the main latency script is unaffected since we're a separate process).
# =============================================================================
while ($true) {
    try {
        $clientBatches = @(Get-ChildItem -Path $NvidiaRootPath -Directory -Filter "batch*" -ErrorAction SilentlyContinue | Sort-Object Name)

        foreach ($dir in $clientBatches) {
            if ($null -eq $dir -or $null -eq $dir.Name -or $null -eq $dir.FullName) { continue }
            $batchName = $dir.Name

            # Derive queue key from folder content timestamps
            # If batch folder is repopulated by a new session, content timestamps change
            # and we get a new key, so it's treated as a fresh batch
            $queueKey = Get-BatchQueueKey -BatchPath $dir.FullName
            if ($null -eq $queueKey) {
                Write-Log ("Batch folder disappeared before key could be derived: {0} - skipping" -f $dir.FullName) "WARN"
                continue
            }

            if ($script:ProcessedQueueKeys.ContainsKey($queueKey)) { continue }

            # Check desktop side is reachable before settling wait
            $desktopPath = "\\$TargetHost\C`$\ProgramData\NVIDIA Corporation\nVector\$batchName"
            if (-not (Test-Path $desktopPath)) {
                Write-Log ("Desktop path not yet reachable for {0}: {1} - will retry next poll" -f $batchName, $desktopPath) "WARN"
                continue
            }

            Write-Log ("New batch detected: {0} [{1}] - settling {2}s..." -f $batchName, $queueKey, $SsimBatchSettleSeconds)
            Start-Sleep -Seconds $SsimBatchSettleSeconds

            # Re-derive key after settle in case agent was still writing
            $queueKey = Get-BatchQueueKey -BatchPath $dir.FullName
            if ($null -eq $queueKey) {
                Write-Log ("Batch folder disappeared during settle: {0} - skipping" -f $dir.FullName) "WARN"
                continue
            }

            if ($script:ProcessedQueueKeys.ContainsKey($queueKey)) {
                Write-Log ("Queue key changed during settle; {0} already queued - skipping" -f $queueKey)
                continue
            }

            $entry = Copy-BatchToQueue -BatchName $batchName -QueueKey $queueKey
            if ($null -eq $entry) {
                Write-Log ("Copy failed for {0} - will not mark as processed; will retry next poll" -f $batchName) "WARN"
                continue
            }

            # Mark as queued before running ssim-tool so a recheck during processing doesn't double-queue
            $script:ProcessedQueueKeys[$queueKey] = $entry.ClientQueue

            $result = Invoke-SsimToolOnEntry -Entry $entry
            if ($null -ne $result) {
                Upload-SsimResult -Result $result | Out-Null
            }
        }
    } catch {
        Write-Log ("Unhandled error in daemon loop: {0} | ScriptStackTrace: {1}" -f $_, $_.ScriptStackTrace) "ERROR"
    }

    Start-Sleep -Seconds $SsimDaemonPollIntervalSeconds
}
