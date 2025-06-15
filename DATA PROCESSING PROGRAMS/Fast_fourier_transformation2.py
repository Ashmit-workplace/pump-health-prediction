import os
import pandas as pd
import numpy as np
from scipy.fft import fft, fftfreq
from scipy.signal import detrend

# === CONFIGURATION ===
bands = [(0, 1), (1, 3), (3, 5), (5, 10)]
axes = ['x_smooth', 'y_smooth', 'z_smooth']  # Use smoothed columns

def fft_features(signal, fs):
    if len(signal) < 8:
        return {
            'total_power': 'insufficient data',
            'spectral_centroid': 'insufficient data',
            **{f'band_{lo}_{hi}Hz': 'insufficient data' for lo, hi in bands}
        }

    signal = detrend(signal)
    N = len(signal)
    freqs = fftfreq(N, d=1/fs)
    power = np.abs(fft(signal))**2
    freqs, power = freqs[freqs > 0], power[freqs > 0]
    total_power = np.sum(power)
    centroid = np.sum(freqs * power) / total_power if total_power > 0 else 0

    features = {
        'total_power': total_power,
        'spectral_centroid': centroid
    }
    for lo, hi in bands:
        features[f'band_{lo}_{hi}Hz'] = np.sum(power[(freqs >= lo) & (freqs < hi)])
    return features

def process_fft_file(input_path):
    try:
        df = pd.read_excel(input_path)
        df['datetime'] = pd.to_datetime(df['datetime'], errors='coerce')
        df.dropna(subset=['datetime'], inplace=True)
        df['second'] = df['datetime'].dt.floor('1s')
        df.sort_values('datetime', inplace=True)

        time_deltas = df['datetime'].diff().dt.total_seconds().dropna()
        fs = 1 / time_deltas.mean() if not time_deltas.empty else 100.0

        records = []
        for second, group in df.groupby('second'):
            row = {'datetime': second}
            for axis in axes:
                if axis in group.columns:
                    stats = fft_features(group[axis].dropna().values, fs)
                    row.update({f'{axis}_{k}': v for k, v in stats.items()})
            records.append(row)

        output_path = input_path.replace('.xlsx', '_fft_features_cleaned.xlsx')
        pd.DataFrame(records).to_excel(output_path, index=False)
        print(f"✅ Saved FFT features to: {output_path}")

    except Exception as e:
        print(f"❌ Error processing {input_path}: {e}")

def recursive_fft_analysis(root_folder):
    for dirpath, _, filenames in os.walk(root_folder):
        for file in filenames:
            if (
                file.endswith('_denoised.xlsx') and
                not file.endswith('_fft_features_cleaned.xlsx')
            ):
                full_path = os.path.join(dirpath, file)
                process_fft_file(full_path)

# === USAGE ===
# Replace with your actual denoised data root folder
root_dir = r"D:\extracted data from JSON file ISI\extracted data\data_processing_denoised\sensor_data_cleaned"
recursive_fft_analysis(root_dir)
