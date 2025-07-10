import os
import pandas as pd

def convert_timestamps_in_excels(root_folder):
    for dirpath, _, filenames in os.walk(root_folder):
        for file in filenames:
            if file.endswith('.xlsx') and not file.startswith('~$'):
                file_path = os.path.join(dirpath, file)
                try:
                    df = pd.read_excel(file_path)

                    # Check if 'timestamp' column exists
                    if 'timestamp' in df.columns:
                        df['datetime'] = pd.to_datetime(df['timestamp'], unit='ms')
                        df.to_excel(file_path, index=False)
                        print(f"[✔] Updated: {file_path}")
                    else:
                        print(f"[!] No 'timestamp' column in: {file_path}")
                except Exception as e:
                    print(f"[✘] Error in {file_path}: {e}")


convert_timestamps_in_excels(r"D:\extracted data from JSON file ISI\rerport writing data")
                             