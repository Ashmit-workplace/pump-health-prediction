import os
import pandas as pd
import numpy as np
from scipy.stats import kurtosis
from tqdm import tqdm

def compute_forceful_kurtosis(group):
    if len(group) < 4:
        return pd.Series({
            'second': group['second'].iloc[0],
            'x_kurtosis': 'insufficient data',
            'y_kurtosis': 'insufficient data',
            'z_kurtosis': 'insufficient data'
        })
    return pd.Series({
        'second': group['second'].iloc[0],
        'x_kurtosis': kurtosis(group['x'], fisher=False, bias=False),
        'y_kurtosis': kurtosis(group['y'], fisher=False, bias=False),
        'z_kurtosis': kurtosis(group['z'], fisher=False, bias=False)
    })

def process_force_kurtosis(file_path):
    try:
        df = pd.read_excel(file_path)
        df['datetime'] = pd.to_datetime(df['timestamp'], unit='ms', errors='coerce')
        df.dropna(subset=['datetime'], inplace=True)
        df['second'] = df['datetime'].dt.floor('s')

        kurt_df = df.groupby('second').apply(compute_forceful_kurtosis).reset_index(drop=True)
        
        # Drop original datetime with millisecond precision
        kurt_df['second'] = kurt_df['second'].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        output_path = os.path.splitext(file_path)[0] + "_kurtosis_summary.xlsx"
        kurt_df.to_excel(output_path, index=False)
        print(f"[✓] Saved: {output_path}")
    except Exception as e:
        print(f"[✗] Error in {file_path}: {e}")

def recursive_force_kurtosis(root_folder):
    for dirpath, _, filenames in os.walk(root_folder):
        for file in filenames:
            if file.endswith(".xlsx") and not file.startswith("~$") and "_kurtosis" not in file:
                file_path = os.path.join(dirpath, file)
                process_force_kurtosis(file_path)

if __name__ == "__main__":
    base_folder = r"D:\extracted data from JSON file ISI\extracted data\Time_series_analysis\sensor_data_Kurtosis1"
    recursive_force_kurtosis(base_folder)
