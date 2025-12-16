# Changelog
## [1.1.0] - 2025-12-16
### ✨ Added
v1.1.0 – DV-safe merge with sanitation and audit logging

This release strengthens Excel merge safety by introducing sanitized, SHA-256–based
content signatures for DV-only records. It prevents silent data corruption caused by
duplicate or mistyped DV numbers and adds a full audit trail via an append-only
MERGE_LOG worksheet.

All merge decisions are now traceable, deterministic, and safe for accounting
and LGU-grade workflows.

## [1.0.0] - 2025-12-15
### ✨ Added
[200~feat: add Excel merge & sync script with composite-key deduplication

- Introduced a Python script to safely merge daily Excel extracts into an official master file
- Implemented composite key logic using CheckNo + RefNo to prevent duplicate records
- Added detection and skipping of duplicates within the same incoming extract
- Preserved manual/protected columns while allowing controlled updates to system-managed fields
- Added automatic timestamped backup of the official Excel file before merge
- Included clear console summary for added, updated, and skipped records
- Designed for double-click execution with no user interaction required

Version: v1.0.0 (initial stable release)
