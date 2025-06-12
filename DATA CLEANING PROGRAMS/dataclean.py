import os

def delete_updated_excel_files(root_folder):
    deleted_count = 0
    for dirpath, _, filenames in os.walk(root_folder):
        for file in filenames:
            if file.endswith('_updated.xlsx'):
                full_path = os.path.join(dirpath, file)
                try:
                    os.remove(full_path)
                    print(f"[🗑️] Deleted: {full_path}")
                    deleted_count += 1
                except Exception as e:
                    print(f"[✘] Could not delete {full_path}: {e}")
    print(f"\n✅ Done. Total deleted: {deleted_count} file(s).")

# 🔧 USAGE: Replace with your path
delete_updated_excel_files(r"D:\extracted data from JSON file ISI\extracted data\sensor_data")

