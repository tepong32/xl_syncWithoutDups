📍 Suggested v1.1.0 Milestones (Roadmap)

These are natural, incremental upgrades based on what you already built — no scope creep.

🔹 v1.1.0 – Robustness & Transparency

Add MERGE_LOG worksheet:

Timestamp

Key

Action taken (ADDED / UPDATED / SKIPPED)

Reason for skip (missing key / duplicate)

Log duplicate rows found within the incoming extract

Log rows skipped due to missing key values

🔹 v1.2.0 – Configurable Merge Rules

External config file (JSON or YAML) for:

Sheet name

Key columns

Updateable columns

Protected/manual columns

Allow column mapping by header name instead of hardcoded letters

🔹 v1.3.0 – User-Friendly Execution

Tkinter wrapper:

File picker for master & extract

“Merge Now” button

Status summary window

Optional dry-run mode (no file saved)

🔹 v2.0.0 – Audit-Grade Sync

Change detection (update only if values differ)

Hash-based row comparison

Version stamp inside Excel (merged-by / merged-on)

Optional PDF or Excel summary export