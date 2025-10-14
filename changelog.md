# Changelog

All notable changes to this project are documented in this file following [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) conventions.

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
