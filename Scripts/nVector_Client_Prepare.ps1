<#
nVector_Client_Prepare.ps1
version 2.5.0

PREVIEW / INTERIM RELEASE
This script is part of the Login Enterprise + NVIDIA nVector integration
preview. It is functional and customer-usable but not yet an official
Login Enterprise built-in feature. For more information refer to the
integration repo at github.com/LoginVSI/nVector and the Login Enterprise
documentation at docs.loginvsi.com.

Change log (high level):
  1.0.0 - original combined uploader logic
  1.0.1 - separated server/local drift helper; initial tests
  1.0.2 - introduced RTT / RawUtc / WallClock modes
  1.0.3 - normalize +0000 -> +00:00 handling
  1.0.4 - added SanityMaxHours + improved logging
  1.0.5 - WallClock default + ForceLocalOffset option
  1.0.6 - WallClock parsing more tolerant (TryParseExact + fallback to Parse)
  1.0.7 - flip parsing order: Try Parse() first, TryParseExact loop as fallback
  2.0.0 - agent 2.0 long-form flags; SSIM integration with batch folder detection;
          ssim-tool invocation and best-score upload to LE API; $SeparateSsimEnvironment
          toggle for multi-y-axis workaround; $HeatmapFolder optional; $MinLatencyThreshold
          added; all paths to C:\ProgramData\NVIDIA Corporation\nVector\;
          retry logic on API uploads; CLI parameter overrides for all key vars; preview banner
  2.1.0 - window title trigger mode: -TargetWindowTitleMatch watches for a remote/console
          window before starting nvector-agent; kills agent and runs SSIM when window closes;
          restarts on next window open; -Mode param: latency | ssim | both (default)
  2.1.1 - promote hardcoded timing values to named config vars
  2.1.2 - add $MinSsimThreshold config var
  2.1.3 - fix ssim-tool argument path quoting; remove in-session SSIM batch detection
  2.1.4 - timestamp every log line via Write-Host wrapper
  2.1.5 - add -NvectorScreenshotIntervalOverride CLI param
  2.1.6 - add CLI override params for all remaining config vars
  2.1.7 - add Invoke-PendingLatency: drain remaining CSV rows on session end
  2.1.8 - ssim-tool invoked once with root paths (reverted in 2.1.9)
  2.1.9 - revert to per-batch ssim-tool invocation; add -SsimBatchCount param
  2.2.0 - add CLI params for WindowPollIntervalMs, ApiRetryCount, ApiRetryDelayMs,
          ApiEndpointLatency, ApiEndpointSsim; add window focus feature (FocusTargetWindow)
  2.3.0 - SSIM processing rewrite: batches are detected and queued as they appear during
          the session (not just at session end); client+desktop screenshots are immediately
          copied to a timestamped ssim_queue subfolder so agent restart cannot clobber
          in-flight ssim-tool runs; ssim-tool is run per queue entry while the session is
          still active; heatmap archiving added: after each batch, heatmaps are copied to
          $HeatmapArchiveFolder\<ts>_<batchName> for permanent retention regardless of
          agent restarts; HWND tracking added: window handle is compared each poll so a
          rapid LE close/reopen (same title, new HWND) is correctly treated as a new
          session; clock adjustment toggle added: -EnableClockAdjustment $false skips
          server time query and uses local clock as-is; per-session summary log written
          at end of each session (session ID, start/end, duration, latency rows uploaded,
          SSIM batches processed); -SsimBatchCount param removed (queue model replaces it)
  2.4.0 - non-blocking SSIM: ssim-tool processing moved to nVector_SSIM_Worker.ps1,
          spawned as a detached process per batch queue entry; main polling loop no longer
          blocks during ssim-tool execution (was 30-90s per batch); worker handles
          ssim-tool + upload + heatmap archive independently; if session ends while a
          worker is still running, worker finishes and uploads on its own; session summary
          now reports batches queued and workers launched (not batches completed, since
          workers may still be running); worker script must live in same folder as this
          script ($PSScriptRoot); worker log written alongside ssim-tool output CSV
  2.5.0 - architectural split: all SSIM logic removed from this script entirely and
          moved to nVector_SSIM_Daemon.ps1; daemon is spawned once at startup (if mode
          includes SSIM) and runs for the full lifetime of this script; daemon owns batch
          folder detection, queue management, ssim-tool invocation, heatmap archiving,
          and LE upload; this script owns agent lifecycle, window/HWND tracking, and
          latency only; daemon is terminated on graceful shutdown; session summary no
          longer includes SSIM counts (daemon handles its own logging)
#>

# =============================================================================
# CLI PARAMETERS
# =============================================================================
# MANDATORY - no usable default; set in the "User Config" section below
# or pass on the command line:
#   -BaseUrl
#   -ConfigurationAccessToken
#   -EnvironmentId
#   -TargetHost
#
# All other parameters are optional with built-in defaults.
#
# -TargetWindowTitleMatch (optional):
#   Case-insensitive substring match against visible window titles.
#   Script waits for a matching window -> starts agent -> collects data.
#   Window disappears -> kills agent -> drains latency -> processes SSIM queue.
#   HWND is tracked so a rapid LE close/reopen (same title) starts a new session.
#   NOTE (MVP): one matching window per script instance only.
#
# -Mode:
#   "both"    (default) - latency AND SSIM
#   "latency" - latency only; all SSIM skipped
#   "ssim"    - SSIM only; latency collection skipped
#
# -EnableClockAdjustment:
#   $true  (default) - query LE appliance for server time, adjust CSV timestamps
#   $false - skip server time query; use local clock as-is (useful if appliance
#            is unreachable at script start or clocks are already in sync)
# =============================================================================
param(
    # --- Mandatory (set defaults in User Config below, or pass on CLI) ---
    [string]$BaseUrl                          = "",
    [string]$ConfigurationAccessToken         = "",
    [string]$EnvironmentId                    = "",
    [string]$TargetHost                       = "",

    # --- Paths (optional overrides) ---
    [string]$NvidiaRootPath                   = "",   # default: C:\ProgramData\NVIDIA Corporation\nVector
    [string]$NvectorAgentExePath              = "",   # default: $NvidiaRootPath\nvector-agent.exe
    [string]$SsimToolExePath                  = "",   # default: $NvidiaRootPath\ssim-tool.exe
    [string]$HeatmapFolder                    = "",   # default: "" (disabled); set to enable per-batch heatmap PNGs
    [string]$HeatmapArchiveFolder             = "",   # default: $NvidiaRootPath\heatmap_archive

    # --- SSIM environment ---
    [string]$SsimEnvironmentId                = "",   # default set in User Config below
    [bool]$SeparateSsimEnvironment            = $true,

    # --- Session trigger ---
    [string]$TargetWindowTitleMatch           = "",   # "" = continuous mode (no window trigger)
    [string]$Mode                             = "both",

    # --- Clock adjustment ---
    [bool]$EnableClockAdjustment              = $true,

    # --- Agent parameter overrides (0 / -1 = use script default) ---
    [int]$NvectorScreenshotIntervalOverride   = 0,    # 0 = default 24
    [int]$NvectorScreenshotsPerBatchOverride  = 0,    # 0 = default 4  -- MUST match target .cs
    [int]$NvectorSamplePeriodMsOverride       = 0,    # 0 = default 5000
    [int]$NvectorMaxScreenshotRoundsOverride  = -1,   # -1 = default 0 (unlimited)

    # --- Latency threshold overrides ---
    [int]$MaxLatencyThresholdOverride         = 0,    # 0 = default 1500 ms
    [int]$MinLatencyThresholdOverride         = -1,   # -1 = default 0 ms

    # --- SSIM threshold overrides ---
    [double]$MinSsimThresholdOverride         = -1,   # -1 = default 0.0

    # --- Timing overrides ---
    [int]$SsimToolTimeoutSecondsOverride      = 0,    # 0 = default 300 s
    [int]$PollingIntervalOverride             = 0,    # 0 = default 10 s
    [int]$WindowPollIntervalMsOverride        = 0,    # 0 = default 1000 ms
    [int]$SsimBatchSettleSecondsOverride      = 0,    # 0 = default 3 s
    [int]$SsimPollIntervalSecondsOverride     = 0,    # 0 = default 5 s
    [int]$CsvWaitSecondsOverride              = 0,    # 0 = default 10 s
    [int]$CsvReadDelayMsOverride              = 0,    # 0 = default 500 ms

    # --- API overrides ---
    [int]$ApiRetryCountOverride               = 0,    # 0 = default 3
    [int]$ApiRetryDelayMsOverride             = 0,    # 0 = default 2000 ms
    [string]$ApiEndpointLatencyOverride       = "",   # "" = use default
    [string]$ApiEndpointSsimOverride          = "",   # "" = use default

    # --- Window focus ---
    [bool]$FocusTargetWindow                  = $false,
    [int]$FocusIntervalSecondsOverride        = 0     # 0 = default 30 s
)

$ScriptVersion = "2.5.0"

$Mode = $Mode.ToLower().Trim()
if ($Mode -notin @("both","latency","ssim")) {
    Write-Warning ("Invalid -Mode '{0}'. Defaulting to 'both'." -f $Mode)
    $Mode = "both"
}

$UseTitleTrigger = ($TargetWindowTitleMatch -ne "")

# =============================================================================
# USER CONFIG
# Edit this section to match your environment.
# Every value here can be overridden at runtime via the CLI params above.
# =============================================================================

# --- MANDATORY: LE connection ---
# Set these here so your command line stays clean.
# Sensitive values (tokens, GUIDs) are intentionally kept out of the param block
# to avoid exposure in process lists and shell history.
if ($BaseUrl                  -eq "") { $BaseUrl                  = "https://myDomain.LoginEnterprise.com/" }
if ($ConfigurationAccessToken -eq "") { $ConfigurationAccessToken = "YOUR-API-TOKEN-HERE" }
if ($EnvironmentId            -eq "") { $EnvironmentId            = "YOUR-LATENCY-ENV-GUID" }
if ($TargetHost               -eq "") { $TargetHost               = "YOUR-TARGET-HOSTNAME" }

# --- SSIM environment ---
# SeparateSsimEnvironment = $true (default, RECOMMENDED):
#   SSIM uploads to a separate LE environment ($SsimEnvironmentId).
#   Required workaround for the LE UI single-y-axis limitation per environment.
#   View latency and SSIM side-by-side in two browser tabs.
# SeparateSsimEnvironment = $false:
#   SSIM uploads to the same environment as latency. Simpler but only one
#   metric visible at a time in the LE UI.
if ($SsimEnvironmentId        -eq "") { $SsimEnvironmentId        = "YOUR-SSIM-ENV-GUID" }

# --- Root path ---
# All nVector artifacts on the launcher live here.
# Update if the launcher user cannot write to C:\ProgramData\.
$NvidiaRootDefault = "C:\ProgramData\NVIDIA Corporation\nVector"
if ($NvidiaRootPath -eq "") { $NvidiaRootPath = $NvidiaRootDefault }

# --- Exe paths ---
# Override only if your executables are not in $NvidiaRootPath.
$NvectorAgentExeDefault = Join-Path $NvidiaRootPath "nvector-agent.exe"
$SsimToolExeDefault     = Join-Path $NvidiaRootPath "ssim-tool.exe"
if ($NvectorAgentExePath -eq "") { $NvectorAgentExePath = $NvectorAgentExeDefault }
if ($SsimToolExePath     -eq "") { $SsimToolExePath     = $SsimToolExeDefault     }

# --- Heatmap output ---
# $HeatmapFolder: live working folder where ssim-tool writes heatmap PNGs.
#   Leave "" to skip heatmap generation entirely.
#   Heatmaps are 3-panel PNGs: desktop | client | difference map.
#   Bright areas on the difference map indicate visual degradation.
# $HeatmapArchiveFolder: after each batch, heatmaps are copied here under a
#   timestamped subfolder (<ts>_<batchName>) for permanent retention.
#   Set to a UNC path to archive across machines.
#   Defaults to $NvidiaRootPath\heatmap_archive if left blank.
$HeatmapFolderDefault = ""
if ($HeatmapFolder -eq "") { $HeatmapFolder = $HeatmapFolderDefault }
if ($HeatmapArchiveFolder -eq "") { $HeatmapArchiveFolder = Join-Path $NvidiaRootPath "heatmap_archive" }

# --- Agent 2.0 parameters ---
# --screenshots-per-batch MUST match the value in NVIDIA nVector Desktop Prepare.cs on the target.
# --max-screenshot-rounds: 0 = unlimited (recommended for continuous LE testing).
$NvectorSamplePeriodMs      = 5000  # watermark rotation period ms       (NVIDIA default: 5000)
$NvectorScreenshotInterval  = 24    # screenshot interval in rotations   (NVIDIA default: 24)
$NvectorScreenshotsPerBatch = 4     # screenshots per batch              (NVIDIA default: 4) -- MUST match target .cs
$NvectorMaxScreenshotRounds = 0     # max rounds; 0 = unlimited

if ($NvectorSamplePeriodMsOverride      -gt 0)  { $NvectorSamplePeriodMs      = $NvectorSamplePeriodMsOverride      }
if ($NvectorScreenshotIntervalOverride  -gt 0)  { $NvectorScreenshotInterval  = $NvectorScreenshotIntervalOverride  }
if ($NvectorScreenshotsPerBatchOverride -gt 0)  { $NvectorScreenshotsPerBatch = $NvectorScreenshotsPerBatchOverride }
if ($NvectorMaxScreenshotRoundsOverride -ge 0)  { $NvectorMaxScreenshotRounds = $NvectorMaxScreenshotRoundsOverride }

# --- Latency thresholds ---
# Readings outside this range are discarded as invalid or spurious.
$MinLatencyThreshold = 0     # ms; <= this is discarded (0 = not a valid VDI measurement)
$MaxLatencyThreshold = 1500  # ms; >= this is discarded (spurious spike)

if ($MinLatencyThresholdOverride -ge 0) { $MinLatencyThreshold = $MinLatencyThresholdOverride }
if ($MaxLatencyThresholdOverride -gt 0) { $MaxLatencyThreshold = $MaxLatencyThresholdOverride }

# --- SSIM thresholds ---
$MinSsimThreshold = 0.0  # scores <= this are discarded (0 = processing failure)

if ($MinSsimThresholdOverride -ge 0) { $MinSsimThreshold = $MinSsimThresholdOverride }

# --- Polling timing ---
$PollingInterval         = 10    # s  - main loop cadence (latency poll + SSIM batch watch)
$WindowPollIntervalMs    = 1000  # ms - window title/HWND check cadence
$CsvWaitSeconds          = 10    # s  - how long to wait for perf.csv after agent start
$CsvReadDelayMs          = 500   # ms - pause before each CSV read (lets agent finish writing)
$SsimBatchSettleSeconds  = 3     # s  - wait after batch folder appears before copying to queue
                                 #      (gives agent time to finish writing all images)
$SsimPollIntervalSeconds = 5     # s  - between ssim-tool output CSV polls waiting for 'best' row
$SsimDaemonPollSeconds   = 10    # s  - how often the SSIM daemon scans for new batch folders

if ($PollingIntervalOverride         -gt 0) { $PollingInterval         = $PollingIntervalOverride         }
if ($WindowPollIntervalMsOverride    -gt 0) { $WindowPollIntervalMs    = $WindowPollIntervalMsOverride    }
if ($CsvWaitSecondsOverride          -gt 0) { $CsvWaitSeconds          = $CsvWaitSecondsOverride          }
if ($CsvReadDelayMsOverride          -gt 0) { $CsvReadDelayMs          = $CsvReadDelayMsOverride          }
if ($SsimBatchSettleSecondsOverride  -gt 0) { $SsimBatchSettleSeconds  = $SsimBatchSettleSecondsOverride  }
if ($SsimPollIntervalSecondsOverride -gt 0) { $SsimPollIntervalSeconds = $SsimPollIntervalSecondsOverride }

# --- SSIM tool timeout ---
# How long to wait for ssim-tool to produce a 'best' row before giving up.
# ssim-tool typically takes 30-90s; 5 minutes is safe for slower machines.
$SsimToolTimeoutSeconds = 300
if ($SsimToolTimeoutSecondsOverride -gt 0) { $SsimToolTimeoutSeconds = $SsimToolTimeoutSecondsOverride }

# --- Window focus ---
# When FocusTargetWindow = $true, the matched window is brought to the foreground
# every $FocusIntervalSeconds. Helps ensure GPU rendering stays active.
$FocusIntervalSeconds = 30  # s - only used if -FocusTargetWindow $true
if ($FocusIntervalSecondsOverride -gt 0) { $FocusIntervalSeconds = $FocusIntervalSecondsOverride }

# --- API ---
$ApiEndpointLatency = "publicApi/v8-preview/platform-metrics"
$ApiEndpointSsim    = "publicApi/v8-preview/platform-metrics"
$ApiRetryCount      = 3     # upload attempts before giving up
$ApiRetryDelayMs    = 2000  # ms between retries

if ($ApiEndpointLatencyOverride -ne "") { $ApiEndpointLatency = $ApiEndpointLatencyOverride }
if ($ApiEndpointSsimOverride    -ne "") { $ApiEndpointSsim    = $ApiEndpointSsimOverride    }
if ($ApiRetryCountOverride      -gt 0)  { $ApiRetryCount      = $ApiRetryCountOverride      }
if ($ApiRetryDelayMsOverride    -gt 0)  { $ApiRetryDelayMs    = $ApiRetryDelayMsOverride    }

# --- Metric metadata ---
# Change these if you want different display names / grouping in the LE UI.
$LatencyMetricId      = "nVectorLatencyMetricId"
$LatencyDisplayName   = "Endpoint Latency"
$LatencyUnit          = "Latency (ms)"
$LatencyGroup         = "nVector"
$LatencyComponentType = "vm"

$SsimMetricId         = "nVectorSsimMetricId"
$SsimDisplayName      = "SSIM Score"
$SsimUnit             = "SSIM"
$SsimGroup            = "nVector"
$SsimComponentType    = "vm"

# --- Misc ---
$Instance            = $env:COMPUTERNAME
$LauncherProcessName = "LoginEnterprise.Launcher.UI"
$LauncherExePath     = "C:\Program Files\Login VSI\Login Enterprise Launcher\LoginEnterprise.Launcher.UI.exe"

# --- Clock drift ---
$AdjustmentMode   = "WallClock"  # "WallClock" (default) | "RTT" | "RawUtc"
$ForceLocalOffset = $null        # override local offset e.g. "-07:00"; $null = auto-detect
$SanityMaxHours   = 168          # abort if computed adjustment exceeds this; 0 = disable check

# =============================================================================
# DERIVED PATHS  (do not edit; computed from User Config above)
# =============================================================================
$CsvFilePath          = Join-Path $NvidiaRootPath "perf.csv"
$NvectorScreenshotDir = $NvidiaRootPath   # agent 2.0 creates batch1, batch2... subfolders here
$NvectorLogFile       = Join-Path $NvidiaRootPath "agent.log"
$SsimQueueFolder      = Join-Path $NvidiaRootPath "ssim_queue"

# Daemon script must live in the same folder as this script.
# Spawned once at startup when mode includes SSIM. Never invoked directly by users.
$SsimDaemonScript = Join-Path $PSScriptRoot "nVector_SSIM_Daemon.ps1"

$ScriptTimestamp = (Get-Date).ToString('yyyyMMddTHHmmss')
$TranscriptFile  = Join-Path $NvidiaRootPath "${ScriptTimestamp}_nVector_Client.log"

$EffectiveSsimEnvironmentId = if ($SeparateSsimEnvironment) { $SsimEnvironmentId } else { $EnvironmentId }

# =============================================================================
# PRE-FLIGHT
# =============================================================================
$ErrorActionPreference = 'Continue'
$VerbosePreference     = 'Continue'
$DebugPreference       = 'SilentlyContinue'

# Wrap Write-Host to prepend ISO timestamp to every log line
function Write-Host {
    param(
        [Parameter(Position=0, ValueFromPipeline=$true)]
        [object]$Object = "",
        [switch]$NoNewline,
        [object]$ForegroundColor,
        [object]$BackgroundColor
    )
    $ts  = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    $msg = if ($Object -is [string]) { $Object } else { "$Object" }
    Microsoft.PowerShell.Utility\Write-Host ("[$ts] $msg") -NoNewline:$NoNewline
}

function Ensure-Directory {
    param([string]$Path)
    if ($Path -ne "" -and -not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Host "Created directory: $Path"
    }
}

Ensure-Directory $NvidiaRootPath
Ensure-Directory $SsimQueueFolder
Ensure-Directory $HeatmapArchiveFolder
if ($HeatmapFolder -ne "") { Ensure-Directory $HeatmapFolder }

# =============================================================================
# TRANSCRIPT
# =============================================================================
Start-Transcript -Path $TranscriptFile -Append

Write-Host "============================================================"
Write-Host "nVector Client Prepare - version $ScriptVersion"
Write-Host "PREVIEW / INTERIM RELEASE - Login VSI + NVIDIA nVector"
Write-Host "============================================================"
Write-Host "NvidiaRootPath:           $NvidiaRootPath"
Write-Host "NvectorAgentExePath:      $NvectorAgentExePath"
Write-Host "SsimToolExePath:          $SsimToolExePath"
Write-Host "TargetHost:               $TargetHost"
Write-Host "EnvironmentId (latency):  $EnvironmentId"
Write-Host "SsimEnvironmentId:        $EffectiveSsimEnvironmentId"
Write-Host "SeparateSsimEnvironment:  $SeparateSsimEnvironment"
Write-Host "HeatmapFolder:            $(if ($HeatmapFolder -ne '') { $HeatmapFolder } else { '(disabled)' })"
Write-Host "HeatmapArchiveFolder:     $HeatmapArchiveFolder"
Write-Host "SsimQueueFolder:          $SsimQueueFolder"
Write-Host "SsimDaemonScript:         $SsimDaemonScript$(if (-not (Test-Path $SsimDaemonScript)) { ' *** NOT FOUND ***' })"
Write-Host "MinLatencyThreshold:      $MinLatencyThreshold ms"
Write-Host "MaxLatencyThreshold:      $MaxLatencyThreshold ms"
Write-Host "MinSsimThreshold:         $MinSsimThreshold"
Write-Host "Mode:                     $Mode"
Write-Host "EnableClockAdjustment:    $EnableClockAdjustment"
Write-Host "TargetWindowTitleMatch:   $(if ($UseTitleTrigger) { $TargetWindowTitleMatch } else { '(not set - continuous mode)' })"
Write-Host "FocusTargetWindow:        $FocusTargetWindow$(if ($FocusTargetWindow) { ' (every ' + $FocusIntervalSeconds + 's)' })"
Write-Host "CsvWaitSeconds:           $CsvWaitSeconds"
Write-Host "CsvReadDelayMs:           $CsvReadDelayMs"
Write-Host "============================================================"

# =============================================================================
# AGENT ARGUMENTS
# =============================================================================
$AgentArguments = @(
    "--role",                  "client",
    "--metrics-file",          "`"$CsvFilePath`"",
    "--screenshots-dir",       "`"$NvectorScreenshotDir`"",
    "--log-file",              "`"$NvectorLogFile`"",
    "--sample-period-ms",      "$NvectorSamplePeriodMs",
    "--screenshot-interval",   "$NvectorScreenshotInterval",
    "--screenshots-per-batch", "$NvectorScreenshotsPerBatch",
    "--max-screenshot-rounds", "$NvectorMaxScreenshotRounds"
)

# =============================================================================
# WIN32: WINDOW DETECTION + HWND TRACKING + FOCUS
# HWND tracking lets us detect a new session even when the window title stays
# the same (e.g. rapid LE continuous test close/reopen).
# =============================================================================
Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public class WindowHelper {
    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    // Returns HWND as Int64 for the first visible window whose title contains
    // titleMatch (case-insensitive). Returns 0 if not found.
    public static long FindWindowHandle(string titleMatch) {
        long found = 0;
        EnumWindows((hWnd, lParam) => {
            if (IsWindowVisible(hWnd)) {
                var sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, 256);
                if (sb.Length > 0 &&
                    sb.ToString().IndexOf(titleMatch, StringComparison.OrdinalIgnoreCase) >= 0) {
                    found = hWnd.ToInt64();
                    return false; // stop enumerating
                }
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static bool FocusWindowByTitle(string titleMatch) {
        bool focused = false;
        EnumWindows((hWnd, lParam) => {
            if (IsWindowVisible(hWnd)) {
                var sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, 256);
                if (sb.Length > 0 &&
                    sb.ToString().IndexOf(titleMatch, StringComparison.OrdinalIgnoreCase) >= 0) {
                    SetForegroundWindow(hWnd);
                    focused = true;
                    return false;
                }
            }
            return true;
        }, IntPtr.Zero);
        return focused;
    }
}
"@

function Get-TargetWindowHandle {
    param([string]$TitleMatch)
    return [WindowHelper]::FindWindowHandle($TitleMatch)
}

function Test-WindowTitleMatch {
    param([string]$TitleMatch)
    return ([WindowHelper]::FindWindowHandle($TitleMatch) -ne 0)
}

function Set-WindowFocus {
    param([string]$TitleMatch)
    if ([WindowHelper]::FocusWindowByTitle($TitleMatch)) {
        Write-Host ("Focused window matching '{0}'" -f $TitleMatch)
    } else {
        Write-Host ("Could not find window matching '{0}' to focus" -f $TitleMatch)
    }
}

# =============================================================================
# CLOCK DRIFT
# =============================================================================
function Get-EffectiveLocalDto {
    if ($null -ne $ForceLocalOffset -and $ForceLocalOffset -ne "") {
        try   { $ts = [TimeSpan]::Parse($ForceLocalOffset) }
        catch { Write-Error ("Invalid ForceLocalOffset '{0}'" -f $ForceLocalOffset); throw }
        return [DateTimeOffset]::UtcNow.ToOffset($ts)
    }
    return [DateTimeOffset]::Now
}

function Compute-AdjustmentCandidates {
    param([string]$BaseUrl, [string]$Endpoint, [string]$Token, [int]$TimeoutSec = 5)
    $uri = $BaseUrl.TrimEnd('/') + '/' + $Endpoint.TrimStart('/')
    $hdr = @{ Authorization = "Bearer $Token" }
    try {
        Write-Host ("Querying server time at {0} ..." -f $uri)
        try   { $r1 = Invoke-WebRequest -Uri $uri -Method Head -Headers $hdr -UseBasicParsing -TimeoutSec $TimeoutSec }
        catch { Write-Host "HEAD failed, falling back to GET..."; $r1 = Invoke-WebRequest -Uri $uri -Method Get -Headers $hdr -UseBasicParsing -TimeoutSec $TimeoutSec }
        $date1 = $r1.Headers['Date']
        if (-not $date1) { throw "No Date header (initial request)" }
        $serverDto1      = [DateTimeOffset]::Parse($date1)
        $serverUtc1      = $serverDto1.UtcDateTime
        $localDto        = Get-EffectiveLocalDto
        $localUtc        = $localDto.UtcDateTime
        $adjustRawUtc    = $serverUtc1 - $localUtc
        $serverAsLocal   = $serverDto1.ToOffset($localDto.Offset)
        $adjustWallClock = $serverAsLocal - $localDto
        Write-Host ("Server UTC: {0}" -f $serverUtc1.ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'"))
        Write-Host ("Local  UTC: {0}" -f $localUtc.ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'"))
        $localBefore = [DateTimeOffset]::UtcNow
        try   { $r2 = Invoke-WebRequest -Uri $uri -Method Head -Headers $hdr -UseBasicParsing -TimeoutSec $TimeoutSec }
        catch { Write-Host "HEAD drift request failed, falling back to GET..."; $r2 = Invoke-WebRequest -Uri $uri -Method Get -Headers $hdr -UseBasicParsing -TimeoutSec $TimeoutSec }
        $localAfter        = [DateTimeOffset]::UtcNow
        $date2             = $r2.Headers['Date']
        if (-not $date2)   { throw "No Date header (drift request)" }
        $serverDto2        = [DateTimeOffset]::Parse($date2)
        $serverUtc2        = $serverDto2.UtcDateTime
        $roundtrip         = $localAfter - $localBefore
        $halfRt            = [TimeSpan]::FromTicks([Math]::Floor($roundtrip.Ticks / 2))
        $estimatedLocalUtc = ($localBefore + $halfRt).UtcDateTime
        $adjustRtt         = $serverUtc2 - $estimatedLocalUtc
        return [PSCustomObject]@{
            AdjustRawUtc    = $adjustRawUtc
            AdjustWallClock = $adjustWallClock
            AdjustRtt       = $adjustRtt
        }
    } catch {
        Write-Error ("Compute-AdjustmentCandidates failed: {0}" -f $_.Exception.Message)
        throw
    }
}

if ($EnableClockAdjustment) {
    Write-Host ("Clock adjustment mode: {0}" -f $AdjustmentMode)
    $adjInfo = Compute-AdjustmentCandidates -BaseUrl $BaseUrl -Endpoint "v8-preview/system/version" -Token $ConfigurationAccessToken
    if (-not $adjInfo) { Write-Error "Adjustment computation failed"; Stop-Transcript; exit 1 }
    switch ($AdjustmentMode.ToUpperInvariant()) {
        "WALLCLOCK" { $script:AdjustToServer = $adjInfo.AdjustWallClock }
        "RTT"       { $script:AdjustToServer = $adjInfo.AdjustRtt       }
        "RAWUTC"    { $script:AdjustToServer = $adjInfo.AdjustRawUtc    }
        Default     { Write-Host ("Unknown mode '{0}', using WallClock." -f $AdjustmentMode); $script:AdjustToServer = $adjInfo.AdjustWallClock }
    }
    $adjMs = [math]::Round($script:AdjustToServer.TotalMilliseconds)
    Write-Host ("Adjustment: {0} ms ({1})" -f $adjMs, $AdjustmentMode)
    if (($SanityMaxHours -gt 0) -and ([math]::Abs($script:AdjustToServer.TotalHours) -gt $SanityMaxHours)) {
        Write-Error ("Adjustment {0} ms exceeds sanity cap of {1} h. Aborting." -f $adjMs, $SanityMaxHours)
        Stop-Transcript; exit 1
    }
} else {
    Write-Host "Clock adjustment disabled (EnableClockAdjustment=$false) - using local clock as-is."
    $script:AdjustToServer = [TimeSpan]::Zero
}
$script:TimeOffsetSpan = $script:AdjustToServer

# =============================================================================
# TIMESTAMP ADJUSTMENT
# =============================================================================
function Adjust-TimeOffset {
    param([string]$RawTimestamp)
    $raw = $RawTimestamp.Trim()
    if ($AdjustmentMode.ToUpperInvariant() -eq "WALLCLOCK" -or -not $EnableClockAdjustment) {
        $stripped = ($raw -replace '([Zz]|[+-]\d{2}:?\d{2})$','').Trim()
        $formats  = @("yyyy-MM-ddTHH:mm:ss.fff","yyyy-MM-ddTHH:mm:ss","yyyy-MM-dd HH:mm:ss.fff","yyyy-MM-dd HH:mm:ss")
        $culture  = [System.Globalization.CultureInfo]::InvariantCulture
        $styles   = [System.Globalization.DateTimeStyles]::AssumeLocal
        $dtLocal  = $null
        try { $dtLocal = [DateTime]::Parse($stripped, $culture, $styles) }
        catch {
            $ok = $false
            foreach ($fmt in $formats) {
                $ref = New-Object System.Object; $refDt = [ref]$ref
                if ([DateTime]::TryParseExact($stripped, $fmt, $culture, $styles, [ref]$refDt)) { $dtLocal = $refDt.Value; $ok = $true; break }
            }
            if (-not $ok) { Write-Warning ("Failed to parse timestamp: '{0}'" -f $stripped); return $null }
        }
        $lDto = Get-EffectiveLocalDto
        try { $lDtoC = New-Object System.DateTimeOffset ($dtLocal, $lDto.Offset) }
        catch { $lDtoC = New-Object System.DateTimeOffset -ArgumentList ($dtLocal.Year,$dtLocal.Month,$dtLocal.Day,$dtLocal.Hour,$dtLocal.Minute,$dtLocal.Second,$dtLocal.Millisecond,$lDto.Offset) }
        return $lDtoC.Add($script:AdjustToServer).UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'")
    }
    if ($raw -match '([+-]\d{4})$') { $raw = $raw -replace '([+-]\d{2})(\d{2})$','$1:$2' }
    try {
        $dto = [DateTimeOffset]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture)
        return $dto.UtcDateTime.Add($script:AdjustToServer).ToString("yyyy-MM-ddTHH:mm:ss.fff'Z'")
    } catch {
        Write-Warning ("Invalid timestamp: '{0}' - {1}" -f $raw, $_.Exception.Message); return $null
    }
}

# =============================================================================
# AGENT HELPERS
# =============================================================================
function Terminate-NvectorAgent {
    $p = Get-Process -Name "nvector-agent" -ErrorAction SilentlyContinue
    if ($p) { Write-Host "Killing nvector-agent..."; $p | Stop-Process -Force }
    else    { Write-Host "nvector-agent not running." }
}

function Start-NvectorAgent {
    if (-not (Test-Path $NvectorAgentExePath)) {
        Write-Error "nvector-agent.exe not found at: $NvectorAgentExePath"
        Stop-Transcript; exit 1
    }
    Start-Process -FilePath $NvectorAgentExePath -ArgumentList $AgentArguments -NoNewWindow -ErrorAction Stop
    Write-Host "nvector-agent started."
    Write-Host ("Arguments: {0}" -f ($AgentArguments -join " "))
}

# =============================================================================
# API UPLOAD WITH RETRY
# =============================================================================
function Upload-DataToApi {
    param([array]$Metrics, [string]$EndpointPath)
    $json = $Metrics | ConvertTo-Json -Depth 10 -Compress
    if (-not $json.TrimStart().StartsWith("[")) { $json = "[$json]" }
    Write-Host ("Uploading {0} metric(s) to {1}" -f $Metrics.Count, $EndpointPath)
    $hdr = @{ Authorization = "Bearer $ConfigurationAccessToken"; "Content-Type" = "application/json" }
    $url = $BaseUrl.TrimEnd('/') + '/' + $EndpointPath.TrimStart('/')
    for ($attempt = 1; $attempt -le $ApiRetryCount; $attempt++) {
        try {
            Invoke-RestMethod -Uri $url -Method Post -Headers $hdr -Body $json | Out-Null
            Write-Host ("Upload succeeded (attempt {0})" -f $attempt)
            return $true
        } catch {
            Write-Warning ("Upload attempt {0}/{1} failed: {2}" -f $attempt, $ApiRetryCount, $_)
            if ($attempt -lt $ApiRetryCount) { Write-Host ("Retrying in {0} ms..." -f $ApiRetryDelayMs); Start-Sleep -Milliseconds $ApiRetryDelayMs }
        }
    }
    Write-Error ("Upload failed after {0} attempts." -f $ApiRetryCount)
    return $false
}

# =============================================================================
# SSIM DAEMON
# nVector_SSIM_Daemon.ps1 is spawned once at startup if mode includes SSIM.
# It runs for the full lifetime of this script and is terminated on shutdown.
# The daemon owns all batch detection, queue management, ssim-tool invocation,
# heatmap archiving, and LE upload. This script does not touch any of that.
# =============================================================================
$script:SsimDaemonProcess = $null

function Start-SsimDaemon {
    if ($Mode -eq "latency") { return }

    if (-not (Test-Path $SsimDaemonScript)) {
        Write-Warning ("SSIM daemon script not found: {0} - SSIM will not run." -f $SsimDaemonScript)
        return
    }

    $daemonTs  = (Get-Date).ToString('yyyyMMddTHHmmss')
    $daemonLog = Join-Path $NvidiaRootPath "${daemonTs}_ssim_daemon.log"

    $daemonArgs = @(
        "-ExecutionPolicy", "Bypass",
        "-NonInteractive",
        "-File",                          "`"$SsimDaemonScript`"",
        "-NvidiaRootPath",                "`"$NvidiaRootPath`"",
        "-SsimToolExePath",               "`"$SsimToolExePath`"",
        "-TargetHost",                    "`"$TargetHost`"",
        "-HeatmapFolder",                 "`"$HeatmapFolder`"",
        "-HeatmapArchiveFolder",          "`"$HeatmapArchiveFolder`"",
        "-SsimQueueFolder",               "`"$SsimQueueFolder`"",
        "-QueueStartupBehavior",          "wipe",
        "-BaseUrl",                       "`"$BaseUrl`"",
        "-ConfigurationAccessToken",      "`"$ConfigurationAccessToken`"",
        "-SsimEnvironmentId",             "`"$EffectiveSsimEnvironmentId`"",
        "-ApiEndpointSsim",               "`"$ApiEndpointSsim`"",
        "-ApiRetryCount",                 "$ApiRetryCount",
        "-ApiRetryDelayMs",               "$ApiRetryDelayMs",
        "-SsimMetricId",                  "`"$SsimMetricId`"",
        "-SsimDisplayName",               "`"$SsimDisplayName`"",
        "-SsimUnit",                      "`"$SsimUnit`"",
        "-SsimGroup",                     "`"$SsimGroup`"",
        "-SsimComponentType",             "`"$SsimComponentType`"",
        "-Instance",                      "`"$Instance`"",
        "-MinSsimThreshold",              "$MinSsimThreshold",
        "-SsimDaemonPollIntervalSeconds", "$SsimDaemonPollSeconds",
        "-SsimBatchSettleSeconds",        "$SsimBatchSettleSeconds",
        "-SsimToolTimeoutSeconds",        "$SsimToolTimeoutSeconds",
        "-SsimPollIntervalSeconds",       "$SsimPollIntervalSeconds",
        "-DaemonLogFile",                 "`"$daemonLog`""
    )

    Write-Host "Starting SSIM daemon..."
    Write-Host ("  daemon log: {0}" -f $daemonLog)
    try {
        $script:SsimDaemonProcess = Start-Process -FilePath "powershell.exe" `
                                                   -ArgumentList $daemonArgs `
                                                   -WindowStyle Hidden `
                                                   -PassThru `
                                                   -ErrorAction Stop
        Write-Host ("SSIM daemon started (PID: {0})." -f $script:SsimDaemonProcess.Id)
    } catch {
        Write-Warning ("Failed to start SSIM daemon: {0}" -f $_.Exception.Message)
        $script:SsimDaemonProcess = $null
    }
}

function Stop-SsimDaemon {
    if ($null -eq $script:SsimDaemonProcess) { return }
    if ($script:SsimDaemonProcess.HasExited) {
        Write-Host ("SSIM daemon already exited (PID: {0})." -f $script:SsimDaemonProcess.Id)
        return
    }
    Write-Host ("Stopping SSIM daemon (PID: {0})..." -f $script:SsimDaemonProcess.Id)
    try {
        $script:SsimDaemonProcess | Stop-Process -Force -ErrorAction SilentlyContinue
        Write-Host "SSIM daemon stopped."
    } catch {
        Write-Warning ("Could not stop SSIM daemon: {0}" -f $_.Exception.Message)
    }
    $script:SsimDaemonProcess = $null
}

# =============================================================================
# LATENCY HELPERS
# =============================================================================
function Invoke-PollingIteration {
    param([ref]$LastLatencyLine)
    if ($Mode -eq "ssim") { return }
    if (-not (Test-Path $CsvFilePath)) { return }
    Start-Sleep -Milliseconds $CsvReadDelayMs
    $all = Get-Content $CsvFilePath
    if ($all.Count -le 1) { return }
    $current = $all.Count - 1
    if ($current -le $LastLatencyLine.Value) { return }
    $newLines = $all[($LastLatencyLine.Value + 1)..$current]
    $metrics  = @()
    foreach ($line in $newLines) {
        $parts = $line -split ','
        if ($parts.Count -ne 2) { Write-Host ("Malformed CSV line: {0}" -f $line); continue }
        $tsRaw = $parts[0].Trim(); $lat = $parts[1].Trim()
        [double]$val = 0.0
        if (-not [double]::TryParse($lat, [ref]$val)) { Write-Host ("Non-numeric latency: '{0}'" -f $lat); continue }
        if ($val -le $MinLatencyThreshold) { Write-Host ("Latency {0}ms <= min {1}ms - discarding" -f $val, $MinLatencyThreshold); continue }
        if ($val -ge $MaxLatencyThreshold) { Write-Host ("Latency {0}ms >= max {1}ms - discarding" -f $val, $MaxLatencyThreshold); continue }
        $ts = Adjust-TimeOffset -RawTimestamp $tsRaw
        if (-not $ts) { Write-Host ("Timestamp error, skipping: {0}" -f $line); continue }
        $metrics += [PSCustomObject]@{
            metricId       = $LatencyMetricId
            environmentKey = $EnvironmentId
            timestamp      = $ts
            displayName    = $LatencyDisplayName
            unit           = $LatencyUnit
            instance       = $Instance
            value          = $val
            group          = $LatencyGroup
            componentType  = $LatencyComponentType
        }
    }
    if ($metrics.Count -gt 0) { Upload-DataToApi -Metrics $metrics -EndpointPath $ApiEndpointLatency }
    $LastLatencyLine.Value = $current
}

function Invoke-PendingLatency {
    param([ref]$LastLatencyLine)
    if ($Mode -eq "ssim") { return }
    if (-not (Test-Path $CsvFilePath)) { Write-Host "No latency CSV to drain."; return }
    Write-Host "Draining remaining latency rows..."
    Invoke-PollingIteration -LastLatencyLine $LastLatencyLine
    Write-Host "Latency drain complete."
}

# =============================================================================
# SESSION SUMMARY LOG
# =============================================================================
function Write-SessionSummary {
    param(
        [string]$SessionId,
        [datetime]$StartTime,
        [int]$LatencyRowsUploaded
    )
    $endTime  = Get-Date
    $duration = ($endTime - $StartTime).ToString("hh\:mm\:ss")
    $summaryFile = Join-Path $NvidiaRootPath "${SessionId}_session_summary.log"
    @(
        "Session ID:     $SessionId",
        "Start:          $($StartTime.ToString('yyyy-MM-ddTHH:mm:ss'))",
        "End:            $($endTime.ToString('yyyy-MM-ddTHH:mm:ss'))",
        "Duration:       $duration",
        "Mode:           $Mode",
        "Latency rows:   $LatencyRowsUploaded",
        "SSIM:           handled by nVector_SSIM_Daemon.ps1 (see ssim_daemon.log)"
    ) | Out-File -FilePath $summaryFile -Encoding utf8
    Write-Host ("Session summary: {0}" -f $summaryFile)
}

# =============================================================================
# START LE LAUNCHER
# =============================================================================
if (-not (Get-Process -Name $LauncherProcessName -ErrorAction SilentlyContinue)) {
    Start-Process -FilePath $LauncherExePath -NoNewWindow -ErrorAction SilentlyContinue
    Write-Host "LE Launcher started."
} else {
    Write-Host "LE Launcher already running."
}

# =============================================================================
# START SSIM DAEMON
# Spawned once here and runs for the lifetime of this script.
# The daemon watches for batch folders, processes SSIM, and uploads independently.
# Nothing in the main execution loop touches SSIM at all.
# =============================================================================
Start-SsimDaemon

# =============================================================================
# MAIN EXECUTION
#
# A) Title trigger mode (-TargetWindowTitleMatch provided):
#    - Wait for matching window (by HWND)
#    - HWND tracked: same title, new HWND = new session (handles rapid LE close/reopen)
#    - Start agent; collect latency only
#    - Window closes: kill agent, drain remaining latency rows, write session summary
#    - Loop indefinitely for next session
#
# B) Continuous mode (no -TargetWindowTitleMatch):
#    - Start agent immediately; poll latency indefinitely
#
# SSIM is handled entirely by nVector_SSIM_Daemon.ps1 in both modes.
# =============================================================================
try {

    if ($UseTitleTrigger) {

        Write-Host "============================================================"
        Write-Host ("Title trigger mode. Watching for: '{0}'" -f $TargetWindowTitleMatch)
        Write-Host ("Window poll interval: {0} ms" -f $WindowPollIntervalMs)
        Write-Host "============================================================"

        $currentHwnd = 0L   # HWND of the session window currently being monitored

        while ($true) {

            # --- Wait for a new window (or the first window) ---
            Write-Host ("Waiting for window matching '{0}'..." -f $TargetWindowTitleMatch)
            $newHwnd = 0L
            while ($newHwnd -eq 0L -or $newHwnd -eq $currentHwnd) {
                Start-Sleep -Milliseconds $WindowPollIntervalMs
                $newHwnd = Get-TargetWindowHandle -TitleMatch $TargetWindowTitleMatch
            }

            $currentHwnd  = $newHwnd
            $sessionId    = (Get-Date).ToString('yyyyMMddTHHmmss')
            $sessionStart = Get-Date
            Write-Host ("Window detected (HWND: {0}). Session: {1}" -f $currentHwnd, $sessionId)

            # --- Start agent ---
            if ($Mode -ne "ssim") {
                Terminate-NvectorAgent
                Start-NvectorAgent
            }

            # --- Per-session state ---
            $LastLatencyLine    = [ref]0
            $LastFocusTime      = [datetime]::MinValue
            $SessionLatencyRows = 0

            # --- Wait for latency CSV ---
            if ($Mode -ne "ssim") {
                $csvExists = $false
                for ($i = 0; $i -lt $CsvWaitSeconds; $i++) {
                    if (Test-Path $CsvFilePath) { $csvExists = $true; break }
                    Start-Sleep -Seconds 1
                }
                if ($csvExists) {
                    $LastLatencyLine.Value = (Get-Content $CsvFilePath).Count - 1
                    Write-Host "Latency CSV ready."
                } else {
                    Write-Host "Latency CSV not found yet - agent may still be initializing."
                }
            }

            Write-Host "Entering session polling loop..."

            # --- Poll while the SAME window (same HWND) is open ---
            while ((Get-TargetWindowHandle -TitleMatch $TargetWindowTitleMatch) -eq $currentHwnd) {

                # Window focus
                if ($FocusTargetWindow) {
                    if (([datetime]::Now - $LastFocusTime).TotalSeconds -ge $FocusIntervalSeconds) {
                        Set-WindowFocus -TitleMatch $TargetWindowTitleMatch
                        $LastFocusTime = [datetime]::Now
                    }
                }

                # Latency
                $linesBefore = $LastLatencyLine.Value
                Invoke-PollingIteration -LastLatencyLine $LastLatencyLine
                $SessionLatencyRows += ($LastLatencyLine.Value - $linesBefore)

                Start-Sleep -Seconds $PollingInterval
            }

            Write-Host ("Window '{0}' closed/changed - session ended." -f $TargetWindowTitleMatch)

            # --- Session teardown ---
            if ($Mode -ne "ssim") { Terminate-NvectorAgent }

            $linesBefore = $LastLatencyLine.Value
            Invoke-PendingLatency -LastLatencyLine $LastLatencyLine
            $SessionLatencyRows += ($LastLatencyLine.Value - $linesBefore)

            Write-SessionSummary -SessionId $sessionId -StartTime $sessionStart `
                                 -LatencyRowsUploaded $SessionLatencyRows

            # Delete stale perf.csv so next session CSV wait does not find old file
            # and set $LastLatencyLine to the wrong position.
            if (Test-Path $CsvFilePath) {
                try { Remove-Item $CsvFilePath -Force -ErrorAction Stop; Write-Host "Latency CSV cleared for next session." }
                catch { Write-Host ("Warning: could not delete latency CSV: {0}" -f $_.Exception.Message) }
            }

            Write-Host "Session complete. Waiting for next window..."
            Write-Host "------------------------------------------------------------"
            $currentHwnd = 0L
        }

    } else {

        # --- Continuous mode ---
        Write-Host "Continuous mode - starting nvector-agent immediately."
        Terminate-NvectorAgent
        if ($Mode -ne "ssim") { Start-NvectorAgent }
        $LastLatencyLine = [ref]0

        if ($Mode -ne "ssim") {
            $csvExists = $false
            for ($i = 0; $i -lt $CsvWaitSeconds; $i++) {
                if (Test-Path $CsvFilePath) { $csvExists = $true; break }
                Start-Sleep -Seconds 1
            }
            if ($csvExists) {
                $LastLatencyLine.Value = (Get-Content $CsvFilePath).Count - 1
                Write-Host "Latency CSV ready."
            } else {
                Write-Host "Latency CSV not found - agent may still be initializing."
            }
        }

        Write-Host "Entering continuous polling loop..."
        while ($true) {
            Invoke-PollingIteration -LastLatencyLine $LastLatencyLine
            Start-Sleep -Seconds $PollingInterval
        }
    }

} finally {
    Write-Host "============================================================"
    Write-Host "nVector Client Prepare shutting down..."
    Write-Host "============================================================"
    Terminate-NvectorAgent
    Stop-SsimDaemon
    Write-Host "Shutdown complete."
    Stop-Transcript
}
