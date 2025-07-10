import os

# === CONFIGURATION ===
root_folder = r"D:\extracted data from JSON file ISI\FINAL BIG DATA\sensor_data"

# Define unwanted suffix patterns in filenames
delete_suffixes = [
    
    "_mps2.xlsx"
]

deleted_files = []

# === DELETE FILES ===
for dirpath, _, filenames in os.walk(root_folder):
    for file in filenames:
        for suffix in delete_suffixes:
            if file.endswith(suffix):
                file_path = os.path.join(dirpath, file)
                try:
                    os.remove(file_path)
                    deleted_files.append(file_path)
                    print(f"[✓] Deleted: {file_path}")
                except Exception as e:
                    print(f"[✗] Failed to delete {file_path}: {e}")

print(f"\n✅ Deletion complete. {len(deleted_files)} files removed.")
