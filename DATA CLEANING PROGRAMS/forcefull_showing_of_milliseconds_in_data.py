import os
import pandas as pd

def update_excel_timestamps_in_folder(parent_folder, time_column='datetime'):
    for root, _, files in os.walk(parent_folder):
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)
                try:
                    # Read the Excel file
                    df = pd.read_excel(file_path)

                    # Skip if the expected datetime column is missing
                    if time_column not in df.columns:
                        print(f"Skipped (no '{time_column}' column): {file_path}")
                        continue

                    # Convert and format the datetime column
                    df[time_column] = pd.to_datetime(df[time_column], errors='coerce')
                    df[time_column] = df[time_column].dt.strftime('%Y-%m-%d %H:%M:%S.%f')

                    # Create new file name with '_updated'
                    base_name, ext = os.path.splitext(file)
                    updated_file = f"{base_name}_updated{ext}"
                    updated_path = os.path.join(root, updated_file)

                    # Save the modified DataFrame
                    df.to_excel(updated_path, index=False)
                    print(f"✅ Saved: {updated_path}")

                except Exception as e:
                    print(f"❌ Error processing {file_path}: {e}")


parent_folder = r"D:\extracted data from JSON file ISI\extracted data\sensor_data"
update_excel_timestamps_in_folder(parent_folder)
