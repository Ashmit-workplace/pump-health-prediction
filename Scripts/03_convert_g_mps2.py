import os
import pandas as pd
from pathlib import Path

# Gravitational constant
G_TO_MPS2 = 9.80665

def convert_g_to_mps2_in_folder(root_folder, overwrite=True):
    for dirpath, _, filenames in os.walk(root_folder):
        for file in filenames:
            if file.endswith(".xlsx") and not file.startswith("~$"):
                file_path = os.path.join(dirpath, file)
                try:
                    df = pd.read_excel(file_path)

                    # Proceed only if x, y, z columns exist
                    if all(axis in df.columns for axis in ['x', 'y', 'z']):
                        df['x_mps2'] = df['x'] * G_TO_MPS2
                        df['y_mps2'] = df['y'] * G_TO_MPS2
                        df['z_mps2'] = df['z'] * G_TO_MPS2

                        if overwrite:
                            df.to_excel(file_path, index=False)
                            print(f"[✔] Updated (overwritten): {file_path}")
                        else:
                            new_path = Path(file_path).with_stem(Path(file_path).stem + "_mps2")
                            df.to_excel(new_path, index=False)
                            print(f"[✔] Created new file: {new_path}")
                    else:
                        print(f"[!] Skipped (missing x/y/z): {file_path}")

                except Exception as e:
                    print(f"[✘] Error processing {file_path}: {e}")


convert_g_to_mps2_in_folder(r"D:\extracted data from JSON file ISI\rerport writing data", overwrite=False)
