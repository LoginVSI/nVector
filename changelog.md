# Changelog

All notable changes to this project are documented in this file following [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) conventions.

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
