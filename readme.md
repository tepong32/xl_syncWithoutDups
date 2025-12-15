📌 Excel Merge & Sync Script — Initial Release Summary
🔹 Purpose

Automates the safe merging and synchronization of a daily Excel extract into an official/master Excel file.

Designed for frequently updated financial records where manual user-entered data must be preserved.

Intended to be double-click runnable (no UI required).

🔹 Core Functionality

Merges data from daily_extract.xlsx into official.xlsx.

Operates on a specified worksheet (Sheet1 by default).

Uses explicit, deterministic keys to identify unique records.

Prevents duplicate entries while allowing controlled updates.

🔹 Key & Deduplication Logic

Uses a composite key consisting of:

CheckNo (Column B)

RefNo (Column C)

Key format (internal):

CHK:<CheckNo>|REF:<RefNo>


Skips:

Records already existing in official.xlsx

Duplicate records within the same daily extract

Rows missing required key values

🔹 Update & Append Rules

New records:

Appended to the end of the official file

Existing records:

Updated only in allowed/system-managed columns

Protected/manual columns:

Preserved and never overwritten (e.g., Notes column)

🔹 Column Handling

Assumes the following column structure:

A – Date
B – CheckNo
C – RefNo
D – RefNo2
E – Description
F – Gross
G – Net
H – Notes (manual / protected)


Only columns A–G are synced from the extract.

Column H is reserved for manual user input.

🔹 Safety Features

Automatically creates a timestamped backup of official.xlsx before merging.

Non-destructive design: no deletions or reordering of existing data.

Explicit skip logic prevents silent data corruption.

🔹 Execution Behavior

Requires no user interaction once files are prepared.

Console output provides a merge summary:

Number of records added

Number of records updated

Rows skipped due to missing keys

Rows skipped due to duplicate entries in the incoming file

🔹 Intended Usage

Daily or periodic syncing of financial or transactional Excel data.

Suitable for LGU / accounting / audit-sensitive workflows.

Designed as a foundation for future enhancements (logging, UI, config files).

🔹 Current Limitations (by design, for v1)

Sheet name is hardcoded (Sheet1).

Column positions are fixed.

No GUI (script-only execution).

No merge log sheet (console output only).

🔹 Upgrade-Ready Architecture

Can be extended to:

Support fallback keys

Add merge audit logs

Introduce config-driven column mapping

Wrap with a Tkinter UI for non-technical users