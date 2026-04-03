"""
check_bkmvdata.py
-----------------
Minimal check: does TXT.BKMVDATA parse correctly?

Usage:
    python check_bkmvdata.py <path_to_TXT.BKMVDATA>
"""

import sys
import os

if len(sys.argv) < 2:
    print("Usage: python check_bkmvdata.py <path_to_TXT.BKMVDATA>")
    sys.exit(1)

path = sys.argv[1]

# Step 1: confirm file exists
if not os.path.isfile(path):
    print(f"FAIL: file not found: {path}")
    sys.exit(1)

print(f"File found: {path}")

# Step 2-4: open with cp1255, count records
count_110b = 0
count_100b = 0
total_lines = 0

with open(path, encoding="cp1255") as f:
    for line in f:
        stripped = line.strip()
        if not stripped:
            continue
        total_lines += 1
        if stripped.startswith("110B"):
            count_110b += 1
        elif stripped.startswith("100B"):
            count_100b += 1

# Step 5: print results
print(f"Total lines : {total_lines}")
print(f"110B records: {count_110b}")
print(f"100B records: {count_100b}")

# Step 6-7: verdict
if count_110b > 0 and count_100b > 0:
    print("SUCCESS: structure parsed correctly")
else:
    print("FAIL: records not detected")
