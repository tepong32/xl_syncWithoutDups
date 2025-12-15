# Changelog
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
