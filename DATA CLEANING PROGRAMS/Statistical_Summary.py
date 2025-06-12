import os
import pandas as pd
import numpy as np
from scipy.stats import skew
from tqdm import tqdm

# === CONFIGURATION ===
folder_path = r"D:\extracted data from JSON file ISI\extracted data\sensor_data"
vibration_axes = ['x_mps2', 'y_mps2', 'z_mps2']
required_columns = ['datetime'] + vibration_axes
error_log_path = os.path.join(folder_path, "processing_errors_updated2.log")

# === STATS FUNCTION ===
def compute_stats(grouped, axis):
    return grouped[axis].agg([
        ('mean', 'mean'),
        ('std', lambda x: np.std(x, ddof=1)),
        ('rms', lambda x: np.sqrt(np.mean(x**2))),
        ('skew', lambda x: skew(x, bias=False))
    ]).rename(columns=lambda col: f"{axis}_{col}")

# === GATHER FILES RECURSIVELY ===
excel_files = []
for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
            excel_files.append(os.path.join(root, file))

# === PROCESS EACH FILE ===
with open(error_log_path, "w") as log_file:
    for file_path in tqdm(excel_files, desc="Processing Excel files"):
        try:
            df_raw = pd.read_excel(file_path, header=0)

            # Try to locate required columns (even if shuffled)
            col_map = {}
            for col in required_columns:
                match = [real_col for real_col in df_raw.columns if col in str(real_col).strip().lower()]
                if match:
                    col_map[col] = match[0]

            if 'datetime' not in col_map:
                log_file.write(f"{os.path.basename(file_path)}: Missing 'datetime' column\n")
                continue

            df = df_raw[list(col_map.values())].copy()
            df.columns = list(col_map.keys())  # rename to standard

            df['datetime'] = pd.to_datetime(df['datetime'], errors='coerce')
            df.dropna(subset=['datetime'], inplace=True)
            df.sort_values('datetime', inplace=True)
            df.set_index('datetime', inplace=True)

            if not isinstance(df.index, pd.DatetimeIndex):
                raise TypeError("Index is not a DatetimeIndex")

            # Round timestamps to nearest 1 second
            df['datetime_rounded'] = df.index.floor('1s')
            grouped = df.groupby('datetime_rounded')

            # Compute average values
            avg_values = grouped[vibration_axes].mean()

            # Compute all stats
            stats_combined = pd.DataFrame()
            for axis in vibration_axes:
                stats = compute_stats(grouped, axis)
                stats_combined = pd.concat([stats_combined, stats], axis=1)

            # Final combined output
            full_output = pd.concat([avg_values, stats_combined], axis=1)
            full_output.reset_index(inplace=True)
            full_output.rename(columns={'datetime_rounded': 'datetime'}, inplace=True)

            # Format datetime to full precision
            full_output['datetime'] = pd.to_datetime(full_output['datetime'], errors='coerce')
            full_output['datetime'] = full_output['datetime'].dt.strftime('%Y-%m-%d %H:%M:%S.%f')

            # Save result as _updated_2.xlsx in same folder
            updated_name = os.path.splitext(os.path.basename(file_path))[0] + "_updated_2.xlsx"
            output_path = os.path.join(os.path.dirname(file_path), updated_name)
            full_output.to_excel(output_path, index=False)

        except Exception as e:
            log_file.write(f"{os.path.basename(file_path)}: {str(e)}\n")
