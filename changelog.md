# Changelog
All notable changes to this project are documented in this file following [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) conventions.

## 2026-03-13

### Added
- `nVector_SSIM_Daemon.ps1` v1.0.2 — new long-running SSIM processing daemon, replaces per-session worker architecture:
  - Spawned once by `nVector_Client_Prepare.ps1` at startup; runs for the full script lifetime
  - Owns all SSIM batch detection, queue management, ssim-tool invocation, heatmap archiving, and LE API upload
  - Queue key derived from batch folder name + newest file timestamp (second precision) — survives batch folder reuse across sessions without session boundary awareness
  - Startup behavior configurable: wipe (default), archive, or keep existing queue
  - Heatmap archiving to timestamped subfolders under `-HeatmapArchiveFolder`
  - Graceful handling of batch folders deleted mid-loop at session boundaries
  - PS5 process object guard on `HasExited` to prevent unhandled null ref after ssim-tool exit
  - Tested across Scenarios 1-4: latency only, both modes, heatmaps, and rapid close/reopen

### Changed
- `nVector_Client_Prepare.ps1` updated to v2.5.0:
  - Refactored from per-batch worker processes to daemon model — spawns `nVector_SSIM_Daemon.ps1` once at startup
  - Daemon terminated in `finally` block on script shutdown
  - Fixed stale `perf.csv` race condition on session restart: CSV deleted after latency drain so next session's polling loop starts clean
  - HWND tracking confirmed working across rapid close/reopen — different HWND detected, agent restarted cleanly each time
- `nVector_SSIM_Worker.ps1` retired — superseded by daemon architecture

### Fixed
- `get_nVectorMetrics.ps1` v2.0.0 — PS5 compatibility fixes:
  - Replaced em dash character in summary output that caused a parse error on PS5
  - `-EnvironmentIds` now correctly splits comma-separated strings passed from the PS5 command line
  - Completion log message now correctly reflects warning state when no data is returned

## 2026-02-26
### Added
- `Add-SessionMetrics.ps1` v1.0.0 — interactive PerfMon and WMI session metric discovery
  and registration script for Login Enterprise:
  - Searches PerfMon counter sets and WMI classes on the local machine by keyword
  - Numbered list UI with grouped results — select counters and classes by number
  - WMI property sub-selection — pick all or specific properties per class
  - Per-property unit prompting with smart suggestions (ms, %, Frames/sec, Bytes/sec, etc.)
  - Per-property summarizeOperation prompting (avg/sum/max/min/none) with smart defaults
  - Confirmation step before any API calls
  - Creates `PerformanceCounter` or `WmiQuery` metric definitions via LE API v8-preview
  - Group assignment: add to existing group (by ID or fuzzy name search), create new, or skip
  - Certificate handling ported from `get_nVectorMetrics.ps1` — `-ImportServerCert` / `-KeepCert`
  - Timestamped log to `%TEMP%`, retry logic, CLI param overrides
  - PowerShell 5 compatible
  - Tested against LE appliance v8-preview API — PerfMon path, WMI path, create group,
    add to existing group, and skip group flows all confirmed working

## 2026-02-24

### Changed
- `get_nVectorMetrics.ps1` updated to v2.0.0 — merged best of both
  `get_nVectorMetrics.ps1` and `Get-LEPlatformMetrics.ps1` into a single script:
  - Renamed `-ApiAccessToken` to `-LEApiToken` for consistency
  - Added `-EnvironmentIds` array parameter — supports one or multiple environment UUIDs in a single run
  - Added `-LastHours` convenience parameter (default: 1) — no manual ISO timestamps required unless desired
  - Added `-OutputDir` parameter — auto-generates timestamped CSV, JSON, and log filenames
  - Updated default API version from `v7-preview` to `v8-preview`
  - Added summary table at end of run showing metric names, units, and data point counts
  - Per-environment error handling — one failing environment does not abort the entire run
  - Added `-IsWarning` log level (yellow) separate from error (red)
  - Output directory pre-creation check
  - Exits with code 1 and usage hint if no environment ID is provided
  - Version bumped to 2.0.0

## 2025-09-08

### Added
- `get_nVectorMetrics.ps1` v1.0.1  
  - New parameters: `-ApiVersion` (defaults to `v7-preview`), `-ImportServerCert`, `-KeepCert`.  
  - Certificate import flow: `Get-RemoteCertificates` (tries `SslStream`, falls back to `HttpWebRequest`/`ServicePoint.Certificate`), `Import-ServerCertificates` (imports leaf + chain into `CurrentUser\Root`), and `Remove-ImportedCertificates` for cleanup.  
  - PowerShell 7 support using `Invoke-RestMethod -SkipCertificateCheck`.

### Changed
- Use `System.Collections.ArrayList` for collecting `X509Certificate2` objects (avoids `op_Addition` issues).  
- Build request URL with `System.UriBuilder` for safe escaping and paths.  
- Force TLS 1.2 on PowerShell 5; read streams and export CSV with UTF-8.  
- Make `Write-Log` non-terminating with improved error handling.

### Fixed
- Resolved certificate concatenation errors and other edge-case failures.  
- Prevent terminating errors when writing logs; hardened error handling during certificate retrieval/import and GET requests.

### Security
- Importing certificates into `CurrentUser\Root` and disabling validation are **insecure**; use only in trusted/test environments for troubleshooting.

## 2025-08-04

### Added
- Dynamic server-based clock drift synchronization with RTT compensation in `nVector_Client_Prepare.ps1`.
- Full-session PowerShell transcript for all console output.
- Per-datapoint timestamp adjustment logging via `Write-Host` (raw → adjusted).

### Changed
- Removed `$ScriptLogFile` and the custom `Log` function in favor of console-only logging.
- Replaced static `$TimeOffset` with a real-time clock-drift calculation against the server’s `Date` header.
- Fail-fast behavior (`exit 1`) on time-sync errors.

### Fixed
- Corrected pipeline pollution in `Adjust-TimeOffset` to ensure valid JSON payloads.