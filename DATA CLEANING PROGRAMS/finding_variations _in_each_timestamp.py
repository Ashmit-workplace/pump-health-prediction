import os
import pandas as pd
import numpy as np
from pathlib import Path
from scipy.stats import kurtosis

# RMS function
def rms(series):
    return np.sqrt(np.mean(series ** 2))

# Process a single file
def process_file(file_path):
    try:
        df = pd.read_excel(file_path)

        required_cols = ['datetime', 'x_mps2', 'y_mps2', 'z_mps2']
        if not all(col in df.columns for col in required_cols):
            print(f"[!] Skipping {file_path} — Missing required columns.")
            return

        grouped = df.groupby('datetime')
        single_row_timestamps = 0

        # Store timestamp sizes for logging
        group_sizes = grouped.size()
        single_row_timestamps = (group_sizes == 1).sum()

        features = grouped.agg({
            'x_mps2': [rms, np.std, lambda x: x.max() - x.min(), np.max, np.min, kurtosis],
            'y_mps2': [rms, np.std, lambda x: x.max() - x.min(), np.max, np.min, kurtosis],
            'z_mps2': [rms, np.std, lambda x: x.max() - x.min(), np.max, np.min, kurtosis],
        })

        # Rename columns
        features.columns = [
            'x_rms', 'x_std', 'x_p2p', 'x_max', 'x_min', 'x_kurtosis',
            'y_rms', 'y_std', 'y_p2p', 'y_max', 'y_min', 'y_kurtosis',
            'z_rms', 'z_std', 'z_p2p', 'z_max', 'z_min', 'z_kurtosis'
        ]
        features = features.reset_index()

        # Fill NaN with 0 for std, p2p, kurtosis where needed
        features.fillna(0, inplace=True)

        # Save new file
        updated_path = Path(file_path).with_stem(Path(file_path).stem + "_updated")
        features.to_excel(updated_path, index=False)

        print(f"[✔] Processed: {updated_path.name} | Single-row timestamps: {single_row_timestamps}")

    except Exception as e:
        print(f"[✘] Error processing {file_path}: {e}")

# Recursively process a folder
def process_folder(root_folder):
    for dirpath, _, filenames in os.walk(root_folder):
        for file in filenames:
            if file.endswith('.xlsx') and not file.startswith("~$") and "_updated" not in file:
                full_path = os.path.join(dirpath, file)
                process_file(full_path)


process_folder(r"D:\extracted data from JSON file ISI\extracted data\sensor_data\Op condition-3\Motor")




