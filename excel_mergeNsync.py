from openpyxl import load_workbook
from datetime import datetime
import shutil

MASTER_FILE = "official.xlsx"
INCOMING_FILE = "daily_extract.xlsx"
SHEET = "Sheet1"

UPDATE_COLS = ["A", "B", "C", "D", "E", "F", "G"]

# Backup
backup = f"official_backup_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
shutil.copy(MASTER_FILE, backup)

wb_m = load_workbook(MASTER_FILE)
wb_i = load_workbook(INCOMING_FILE)

ws_m = wb_m[SHEET]
ws_i = wb_i[SHEET]

def resolve_key(ws, row):
    check_no = ws[f"B{row}"].value
    ref_no = ws[f"C{row}"].value

    if check_no and ref_no:
        return f"CHK:{str(check_no).strip()}|REF:{str(ref_no).strip()}"
    return None

# Build master index
master_index = {}
for r in range(2, ws_m.max_row + 1):
    key = resolve_key(ws_m, r)
    if key:
        master_index[key] = r

# Track incoming duplicates too
incoming_seen = set()

added = updated = skipped = dup_incoming = 0

for r in range(2, ws_i.max_row + 1):
    key = resolve_key(ws_i, r)

    if not key:
        skipped += 1
        continue

    # Skip duplicates inside daily_extract itself
    if key in incoming_seen:
        dup_incoming += 1
        continue
    incoming_seen.add(key)

    if key in master_index:
        target = master_index[key]
        for col in UPDATE_COLS:
            ws_m[f"{col}{target}"].value = ws_i[f"{col}{r}"].value
        updated += 1
    else:
        ws_m.append([ws_i[f"{c}{r}"].value for c in UPDATE_COLS] + [""])
        master_index[key] = ws_m.max_row  # 🔥 critical line
        added += 1

wb_m.save(MASTER_FILE)

print("Merge completed.")
print(f"Added: {added}")
print(f"Updated: {updated}")
print(f"Skipped (no key): {skipped}")
print(f"Skipped (duplicate in incoming): {dup_incoming}")
