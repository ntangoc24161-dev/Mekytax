import json
import os
import re

with open('scan_result.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Let's list all folders and see what companies exist
print("=== LIST OF TOP-LEVEL DIRECTORIES ===")
for k in sorted(data.keys()):
    if data[k]['type'] == 'dir':
        print(f"Folder: {k}")

print("\n=== TOP-LEVEL FILES ===")
for k in sorted(data.keys()):
    if data[k]['type'] == 'file':
        print(f"File: {k}")
