import os
import json
import pandas as pd

def convert_json_to_excel_with_updated_suffix(root_folder):
    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.endswith(".json"):
                json_path = os.path.join(dirpath, filename)

                try:
                    # Load JSON file
                    with open(json_path, 'r') as f:
                        data = json.load(f)

                    # Create DataFrame
                    df = pd.DataFrame(data['CSV'], columns=['timestamp', 'x', 'y', 'z'])

                    # Construct new Excel file name with 'updated' suffix
                    base_name = os.path.splitext(filename)[0]
                    excel_filename = f"{base_name}_updated.xlsx"
                    excel_path = os.path.join(dirpath, excel_filename)

                    # Save to Excel
                    df.to_excel(excel_path, index=False)

                    print(f"✅ Converted: {json_path} → {excel_path}")

                except Exception as e:
                    print(f"❌ Failed to convert {json_path}: {e}")


parent_folder = r"D:\extracted data from JSON file ISI\extracted data\sensor_data"
convert_json_to_excel_with_updated_suffix(parent_folder)

