import os
import pandas as pd

def flag_missing_values(filepath):
    try:
        df = pd.read_excel(filepath)
        
        # Identify axis columns (x/y/z in either format)
        axes = ['x_mps2', 'y_mps2', 'z_mps2']
        if not all(axis in df.columns for axis in axes):
            print(f"⚠️ Skipping {filepath}: Missing expected axis columns.")
            return

        # Add missing value flag
        df['is_missing'] = ((df['x_mps2'] == 0.0) | 
                            (df['y_mps2'] == 0.0) | 
                            (df['z_mps2'] == 0.0)).astype(int)

        # Preserve full datetime if available
        if 'datetime' in df.columns:
            df['datetime'] = pd.to_datetime(df['datetime'], errors='coerce')
            df['datetime'] = df['datetime'].dt.strftime('%Y-%m-%d %H:%M:%S.%f').str[:-3]

        # Save new file
        new_path = filepath.replace('.xlsx', '_flagged_missing.xlsx')
        df.to_excel(new_path, index=False)
        print(f"✅ Saved flagged file: {new_path}")

    except Exception as e:
        print(f"❌ Error processing {filepath}: {e}")

def recursively_flag_folder(folder_path):
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx") and not file.endswith('_flagged_missing.xlsx'):
                full_path = os.path.join(root, file)
                flag_missing_values(full_path)

# === USAGE ===
# Replace with your folder path
folder_path = r"D:\extracted data from JSON file ISI\FINAL BIG DATA\sensor_data"


if __name__ == "__main__":
    recursively_flag_folder(folder_path)
