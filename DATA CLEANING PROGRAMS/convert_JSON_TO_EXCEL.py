import os
import json
import pandas as pd

def convert_all_json_to_excel(root_folder):
    # Walk through all subdirectories and files
    for dirpath, dirnames, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.endswith(".json"):
                json_path = os.path.join(dirpath, filename)
                
                try:
                    with open(json_path, 'r') as f:
                        data = json.load(f)

                    # Convert to DataFrame
                    df = pd.DataFrame(data['CSV'], columns=['timestamp', 'x', 'y', 'z'])

                    # Create Excel path (same name, same folder)
                    excel_filename = os.path.splitext(filename)[0] + '.xlsx'
                    excel_path = os.path.join(dirpath, excel_filename)

                    # Save to Excel
                    df.to_excel(excel_path, index=False)

                    print(f" Converted: {json_path} → {excel_path}")
                except Exception as e:
                    print(f" Error processing {json_path}: {e}")


parent_folder = r"D:\extracted data from JSON file ISI\extracted data\Off condition\Blower"

convert_all_json_to_excel(parent_folder)
