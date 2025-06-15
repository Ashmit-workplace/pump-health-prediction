import os
import pandas as pd
import numpy as np
from scipy.stats import kurtosis
from tqdm import tqdm

def compute_forceful_kurtosis(group):
    if len(group) < 8:
        return pd.Series({
            'second': group['second'].iloc[0],
            'x_kurtosis': 'insufficient data',
            'y_kurtosis': 'insufficient data',
            'z_kurtosis': 'insufficient data'
        })
    return pd.Series({
        'second': group['second'].iloc[0],
        'x_kurtosis': kurtosis(group['x_smooth'], fisher=False, bias=False),
        'y_kurtosis': kurtosis(group['y_smooth'], fisher=False, bias=False),
        'z_kurtosis': kurtosis(group['z_smooth'], fisher=False, bias=False)
    })

def process_force_kurtosis(file_path):
    try:
        df = pd.read_excel(file_path)
        if 'datetime' not in df.columns:
            print(f"[!] Skipped (no datetime): {file_path}")
            return

        df['datetime'] = pd.to_datetime(df['datetime'], errors='coerce')
        df.dropna(subset=['datetime'], inplace=True)
        df['second'] = df['datetime'].dt.floor('s')

        # Compute per-second kurtosis
        kurt_df = df.groupby('second').apply(compute_forceful_kurtosis).reset_index(drop=True)

        # Convert datetime to full-precision string
        kurt_df['second'] = pd.to_datetime(kurt_df['second'], errors='coerce')
        kurt_df['second'] = kurt_df['second'].dt.strftime('%Y-%m-%d %H:%M:%S.%f')

        # Save result
        output_path = file_path.replace('.xlsx', '_kurtosis_updated.xlsx')
        kurt_df.rename(columns={'second': 'datetime'}, inplace=True)
        kurt_df.to_excel(output_path, index=False)
        print(f"[✓] Saved: {output_path}")
    except Exception as e:
        print(f"[✗] Error in {file_path}: {e}")

def recursive_force_kurtosis(root_folder):
    for dirpath, _, filenames in os.walk(root_folder):
        for file in filenames:
            if (
                file.endswith("_denoised.xlsx")
                and not file.endswith("_kurtosis_updated.xlsx")
                and not file.startswith("~$")
            ):
                file_path = os.path.join(dirpath, file)
                process_force_kurtosis(file_path)

if __name__ == "__main__":
    base_folder = r"E:\extracted data from JSON file ISI\extracted data\data_processing_denoised\sensor_data_cleaned"
    recursive_force_kurtosis(base_folder)