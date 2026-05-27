import json
import os

import pandas as pd

from config import SCS_JSON_PATH, SCS_GRANULAR_FILE_PATH

def clean_granular(file):
    try:
        
        df = pd.read_excel(file.stream, engine='openpyxl') 
        # Add a new column 'Comments' initially with 'ERROR'
        df['Comments'] = 'ERROR'

        # Path to the folder containing JSON files
        json_folder_path = os.path.join(os.getcwd(), 'json')

        # Iterate over each row in the DataFrame
        for index, row in df.iterrows():
            component = row['Component']
            container_value = row['Granular Container Value']
            found = False  # Flag to track if Component is found

            # Check if container_value is "[BLANK]" and set 'Comments' to 'Verify'
            if container_value == "[BLANK]":
                df.at[index, 'Comments'] = 'Verify'
                continue

            # Iterate over JSON files
            for filename in os.listdir(SCS_JSON_PATH):
                if filename.endswith('.json'):
                    json_path = os.path.join(SCS_JSON_PATH, filename)
                    with open(json_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        # Iterate over all keys in the JSON
                        for key, value in data.items():
                            if isinstance(value, list):
                                for entry in value:
                                    if 'Component' in entry and entry['Component'] == component:
                                        # Check if Granular Container Value is contained in ContainerValue
                                        if container_value in entry.get("ContainerValue", ""):
                                            df.at[index, 'Comments'] = 'OK'
                                        found = True  # Component found, stop searching
                                        break
                            if found:
                                break
                if found:
                    break

        # Export DataFrame to Excel with the updated 'Comments' column
        df.to_excel(SCS_GRANULAR_FILE_PATH, index=False)

    except Exception as e:
        print(e)