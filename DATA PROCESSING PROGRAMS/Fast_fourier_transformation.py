import os
import pandas as pd
import numpy as np
from scipy.fft import fft, fftfreq
from tqdm import tqdm

# === CONFIG ===
vibration_axes = ['x_mps2', 'y_mps2', 'z_mps2']
base_fft_feature_names = ['peak_freq', 'centroid_freq', 'total_power']
band_limits = [(0, 10), (10, 20)]  # You can add more bands here if your data supports it

# === FFT Feature Extractor ===
def extract_fft_features(signal, sampling_rate):
    N = len(signal)
    if N < 4:
        return {f: 0.0 for f in base_fft_feature_names + [f'band_power_{lo}_{hi}' for lo, hi in band_limits]}

    yf = np.abs(fft(signal))
    xf = fftfreq(N, d=1 / sampling_rate)

    pos = xf >= 0
    xf = xf[pos]
    yf = yf[pos]
    power = yf ** 2

    features = {}

    total_power = np.sum(power)
    centroid = 0.0 if total_power == 0 else np.sum(xf * power) / total_power
    peak_freq = xf[np.argmax(power)]

    features['peak_freq'] = peak_freq
    features['centroid_freq'] = centroid
    features['total_power'] = total_power

    nyquist = sampling_rate / 2
    for lo, hi in band_limits:
        if hi <= nyquist:
            band_power = np.sum(power[(xf >= lo) & (xf < hi)])
            features[f'band_power_{lo}_{hi}'] = band_power
        else:
            # Drop the band if it's outside resolvable range
            features[f'band_power_{lo}_{hi}'] = 0.0

    return features

# === Process Individual File ===
def process_file_fft(file_path):
    try:
        df = pd.read_excel(file_path)
        if 'datetime' not in df.columns or not all(axis in df.columns for axis in vibration_axes):
            return

        df['datetime'] = pd.to_datetime(df['datetime'], errors='coerce')
        df.dropna(subset=['datetime'], inplace=True)
        df.sort_values('datetime', inplace=True)
        df['second'] = df['datetime'].dt.floor('1s')

        time_diffs = df['datetime'].diff().dt.total_seconds().dropna()
        sampling_rate = 1 / time_diffs.mean() if not time_diffs.empty else 100.0

        results = []
        for ts, group in df.groupby('second'):
            row = {'datetime': ts}
            for axis in vibration_axes:
                features = extract_fft_features(group[axis].values, sampling_rate)
                row.update({f'{axis}_{key}': val for key, val in features.items()})
            results.append(row)

        output_df = pd.DataFrame(results)
        output_name = os.path.splitext(os.path.basename(file_path))[0] + "_fft_filtered.xlsx"
        output_path = os.path.join(os.path.dirname(file_path), output_name)
        output_df.to_excel(output_path, index=False)

    except Exception as e:
        print(f"[!] Error processing {file_path}: {e}")

# === Recursively Search and Process All Files ===
def process_all_excel_files(root_folder):
    all_excel_files = []
    for root, _, files in os.walk(root_folder):
        for file in files:
            if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$') and '_fft' not in file:
                all_excel_files.append(os.path.join(root, file))

    for file in tqdm(all_excel_files, desc="Extracting FFT features"):
        process_file_fft(file)

if __name__ == "__main__":
    root_folder = r"D:\extracted data from JSON file ISI\extracted data\sensor_data"
    process_all_excel_files(root_folder)

