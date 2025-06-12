import os

def delete_non_updated_files(root_folder):
    deleted_count = 0
    kept_count = 0

    for dirpath, _, filenames in os.walk(root_folder):
        for file in filenames:
            full_path = os.path.join(dirpath, file)

            # Condition to keep only "_updated.xlsx" files
            if file.endswith('_updated.xlsx'):
                kept_count += 1
                continue

            # Safe to delete
            try:
                os.remove(full_path)
                print(f"[🗑️] Deleted: {full_path}")
                deleted_count += 1
            except Exception as e:
                print(f"[✘] Could not delete {full_path}: {e}")

    print(f"\n✅ Cleanup Complete — Deleted: {deleted_count} | Kept: {kept_count} updated files.")


delete_non_updated_files(r"D:\extracted data from JSON file ISI\extracted data\sensor_data")
