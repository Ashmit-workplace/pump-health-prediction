import os
from pathlib import Path

def delete_old_excel_files(root_folder):
    for dirpath, _, filenames in os.walk(root_folder):
        for file in filenames:
            full_path = os.path.join(dirpath, file)
            if file.endswith(".xlsx") and not file.endswith("_mps2.xlsx") and not file.startswith("~$"):
                try:
                    os.remove(full_path)
                    print(f"[🗑️] Deleted: {full_path}")
                except Exception as e:
                    print(f"[✘] Failed to delete {full_path}: {e}")

# 👉 USAGE
# Run this after conversion is confirmed
delete_old_excel_files(r"D:\extracted data from JSON file ISI\extracted data\sensor_data")
