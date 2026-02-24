import json
from config import SCS_PRODUCT_LINES_PATH


def pl_check(df):
    
    # Validate required columns
    required_columns = ['PL', 'ContainerValue']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns in pl_check: {missing_columns}. Available columns: {df.columns.tolist()}")
    
    # Read JSON data
    with open(SCS_PRODUCT_LINES_PATH, "r") as json_file:  # Server
        # with open('app/core/data/product_lines.json', 'r') as json_file:  # Local
        json_data = json.load(json_file)

    # Iterate through each product line in the JSON
    for pl_info in json_data["ProductLine"]:
        pl = pl_info["PL"]
        container_names = pl_info["ContainerName"]

        # Check if the PL column in the DataFrame matches the current PL from the JSON
        matching_rows = df[df["PL"] == pl]

        # Iterate through ContainerName columns
        for container_name in container_names:
            # Check if the 'ContainerValue' column is empty in the matching rows
            empty_container_rows = matching_rows[matching_rows["ContainerValue"].isna()]
            
            # Update the 'ContainerValue' column for empty container rows
            df.loc[empty_container_rows.index, "ContainerValue"] = "ERROR: Mandatory Container Value"

    return df

def pl_check_granular(df):
    
    # Read JSON data
    with open(SCS_PRODUCT_LINES_PATH, "r") as json_file:  # Server
        # with open('app/core/data/product_lines.json', 'r') as json_file:  # Local
        json_data = json.load(json_file)

    # Iterate through each product line in the JSON
    for pl_info in json_data["ProductLine"]:
        pl = pl_info["PL"]
        container_names = pl_info["ContainerName"]

        # Check if the PL column in the DataFrame matches the current PL from the JSON
        matching_rows = df[df["PL"] == pl]

        # Iterate through ContainerName columns
        for container_name in container_names:
            # Check if the 'ContainerValue' column is empty in the matching rows
            empty_container_rows = matching_rows[matching_rows["ContainerValue"].isna()]
            
            # Update the 'ContainerValue' column for empty container rows
            df.loc[empty_container_rows.index, "ContainerValue"] = "ERROR: Mandatory Container Value"

    return df