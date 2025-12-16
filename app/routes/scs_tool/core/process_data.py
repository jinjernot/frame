import pandas as pd
import json
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import lru_cache

# Global cache for JSON data to avoid repeated file reads
_json_cache = {}

def load_json_with_cache(json_path):
    """
    Load JSON file with caching to avoid repeated reads.
    """
    if json_path not in _json_cache:
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                _json_cache[json_path] = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            _json_cache[json_path] = None
    return _json_cache[json_path]

def clear_json_cache():
    """
    Clear the JSON cache. Call this if memory becomes an issue.
    """
    global _json_cache
    _json_cache.clear()

def process_data(json_path, container_name, df):
    """
    Processes a standard report, checks accuracy, and provides the correct value on error.
    Optimized version with caching and vectorized operations.
    """
    json_data = load_json_with_cache(json_path)
    
    if json_data is None:
        mask = df['ContainerName'] == container_name
        df.loc[mask, 'Accuracy'] = f'ERROR: {container_name} JSON not found or invalid'
        return df

    container_data = json_data.get(container_name, {})
    
    # Create an inverted map for very fast lookups of the correct value.
    # Maps each component to its one correct container value.
    component_to_value_map = {
        component: value
        for value, components in container_data.items()
        for component in components
    }

    # Create a boolean mask for the rows that match the current container
    mask = df['ContainerName'] == container_name
    
    # Only process if there are relevant rows
    if not mask.any():
        return df
    
    # Get indices where mask is True
    relevant_indices = df.index[mask]
    
    # Vectorized lookup using map
    components = df.loc[relevant_indices, 'Component']
    current_values = df.loc[relevant_indices, 'ContainerValue']
    
    # Map components to correct values
    correct_values = components.map(component_to_value_map)
    
    # Create accuracy status and correct value columns
    matches = current_values == correct_values
    component_not_found = correct_values.isna()
    
    # Set accuracy status
    accuracy = pd.Series('', index=relevant_indices)
    accuracy[matches] = f'SCS {container_name} OK'
    accuracy[~matches & ~component_not_found] = f'ERROR: {container_name}'
    accuracy[component_not_found] = f'ERROR: {container_name}'
    
    # Set correct value column
    correct_value_col = pd.Series('', index=relevant_indices)
    correct_value_col[~matches & ~component_not_found] = correct_values[~matches & ~component_not_found]
    correct_value_col[component_not_found] = 'Component Not Found in JSON'
    
    # Update the DataFrame
    df.loc[relevant_indices, 'Accuracy'] = accuracy
    df.loc[relevant_indices, 'Correct Value'] = correct_value_col
    
    return df


def process_data_granular(json_path, container_name, df_g):
    """
    Processes a granular report, checks accuracy, and provides the correct value on error.
    Optimized version with caching and vectorized operations.
    """
    json_data = load_json_with_cache(json_path)
    
    if json_data is None:
        mask = df_g['Granular Container Tag'] == container_name
        df_g.loc[mask, 'Accuracy'] = f'ERROR: {container_name} JSON not found or invalid'
        return df_g

    container_data = json_data.get(container_name, {})
    
    # Create the inverted map for fast lookups.
    component_to_value_map = {
        component: value
        for value, components in container_data.items()
        for component in components
    }
    
    mask = df_g['Granular Container Tag'] == container_name
    
    # Only process if there are relevant rows
    if not mask.any():
        return df_g
    
    # Get indices where mask is True
    relevant_indices = df_g.index[mask]
    
    # Vectorized lookup using map
    components = df_g.loc[relevant_indices, 'Component']
    current_granular_values = df_g.loc[relevant_indices, 'Granular Container Value']
    
    # Map components to correct values
    correct_values = components.map(component_to_value_map)
    
    # Create accuracy status and correct value columns
    matches = current_granular_values == correct_values
    component_not_found = correct_values.isna()
    
    # Set accuracy status
    accuracy = pd.Series('', index=relevant_indices)
    accuracy[matches] = f'SCS {container_name} OK'
    accuracy[~matches & ~component_not_found] = f'ERROR: {container_name}'
    accuracy[component_not_found] = f'ERROR: {container_name}'
    
    # Set correct value column
    correct_value_col = pd.Series('', index=relevant_indices)
    correct_value_col[~matches & ~component_not_found] = correct_values[~matches & ~component_not_found]
    correct_value_col[component_not_found] = 'Component Not Found in JSON'
    
    # Update the DataFrame
    df_g.loc[relevant_indices, 'Accuracy'] = accuracy
    df_g.loc[relevant_indices, 'Correct Value'] = correct_value_col
    
    return df_g


def process_multiple_containers_parallel(df, json_dir, container_col='ContainerName', max_workers=4):
    """
    Process multiple JSON files in parallel for standard reports.
    
    Args:
        df: DataFrame to process
        json_dir: Directory containing JSON files
        container_col: Column name for container identification
        max_workers: Number of parallel workers
    
    Returns:
        Updated DataFrame
    """
    # Get unique container names that exist in the DataFrame
    containers_in_df = df[container_col].unique()
    
    # Get list of JSON files
    json_files = [f for f in os.listdir(json_dir) if f.endswith('.json')]
    
    # Pre-load all JSON files into cache
    for json_file in json_files:
        json_path = os.path.join(json_dir, json_file)
        load_json_with_cache(json_path)
    
    # Create tasks for containers that exist in both DataFrame and JSON files
    tasks = []
    for json_file in json_files:
        container_name = os.path.splitext(json_file)[0]
        if container_name in containers_in_df:
            json_path = os.path.join(json_dir, json_file)
            tasks.append((json_path, container_name))
    
    # Process in parallel
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_data, json_path, container_name, df): container_name 
                   for json_path, container_name in tasks}
        
        for future in as_completed(futures):
            try:
                df = future.result()
            except Exception as e:
                container_name = futures[future]
                print(f"Error processing {container_name}: {e}")
    
    return df


def process_multiple_containers_parallel_granular(df_g, json_dir, container_col='Granular Container Tag', max_workers=4):
    """
    Process multiple JSON files in parallel for granular reports.
    
    Args:
        df_g: DataFrame to process
        json_dir: Directory containing JSON files
        container_col: Column name for container identification
        max_workers: Number of parallel workers
    
    Returns:
        Updated DataFrame
    """
    # Get unique container names that exist in the DataFrame
    containers_in_df = df_g[container_col].unique()
    
    # Get list of JSON files
    json_files = [f for f in os.listdir(json_dir) if f.endswith('.json')]
    
    # Pre-load all JSON files into cache
    for json_file in json_files:
        json_path = os.path.join(json_dir, json_file)
        load_json_with_cache(json_path)
    
    # Create tasks for containers that exist in both DataFrame and JSON files
    tasks = []
    for json_file in json_files:
        container_name = os.path.splitext(json_file)[0]
        if container_name in containers_in_df:
            json_path = os.path.join(json_dir, json_file)
            tasks.append((json_path, container_name))
    
    # Process in parallel
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_data_granular, json_path, container_name, df_g): container_name 
                   for json_path, container_name in tasks}
        
        for future in as_completed(futures):
            try:
                df_g = future.result()
            except Exception as e:
                container_name = futures[future]
                print(f"Error processing {container_name}: {e}")
    
    return df_g
