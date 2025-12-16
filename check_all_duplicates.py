import re
from collections import defaultdict
import json
from pathlib import Path

# Define paths
BASE_DIR = Path(__file__).resolve().parent
DB_FOLDER = BASE_DIR / "app" / "routes" / "scs_tool" / "data" / "db"
DB_GRANULAR_FOLDER = BASE_DIR / "app" / "routes" / "scs_tool" / "data" / "db_granular"

def check_and_fix_file(file_path):
    """Check a single file for duplicate keys and fix if found."""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Find all key-value pairs
    pattern = r'"([^"]+)":\s*\[((?:[^\[\]]*|\[[^\]]*\])*)\]'
    matches = re.findall(pattern, content, re.DOTALL)
    
    if not matches:
        return None
    
    # Get the main key (first match should be the top-level key)
    main_key = matches[0][0] if matches else None
    
    # Group by key
    key_groups = defaultdict(list)
    for key, value_str in matches:
        if key == main_key:
            continue
        components = re.findall(r'"([^"]+)"', value_str)
        key_groups[key].append(components)
    
    # Find duplicates
    duplicates = {k: v for k, v in key_groups.items() if len(v) > 1}
    
    if not duplicates:
        return None
    
    # Merge and create clean JSON
    merged_data = {main_key: {}}
    for key, component_lists in key_groups.items():
        # Combine all components and remove duplicates while preserving order
        all_components = []
        seen = set()
        for comp_list in component_lists:
            for comp in comp_list:
                if comp not in seen:
                    seen.add(comp)
                    all_components.append(comp)
        merged_data[main_key][key] = sorted(all_components)
    
    # Save the fixed file
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(merged_data, f, indent=4, ensure_ascii=False)
    
    return {
        'file': file_path.name,
        'duplicate_keys': len(duplicates),
        'details': duplicates
    }

def process_folder(folder_path):
    """Process all JSON files in a folder."""
    print(f"\n{'='*80}")
    print(f"Checking folder: {folder_path.name}")
    print(f"{'='*80}")
    
    json_files = sorted(folder_path.glob("*.json"))
    print(f"Found {len(json_files)} JSON files")
    
    files_with_duplicates = []
    
    for json_file in json_files:
        result = check_and_fix_file(json_file)
        if result:
            files_with_duplicates.append(result)
            print(f"\n⚠️  {result['file']}: {result['duplicate_keys']} duplicate keys")
            
            # Show first few examples
            for key, occurrences in list(result['details'].items())[:5]:
                total = sum(len(occ) for occ in occurrences)
                print(f"    - '{key[:70]}{'...' if len(key) > 70 else ''}': {len(occurrences)}x, {total} components")
            
            if len(result['details']) > 5:
                print(f"    ... and {len(result['details']) - 5} more duplicate keys")
    
    return files_with_duplicates

print("="*80)
print("CHECKING ALL JSON FILES FOR DUPLICATE KEYS")
print("="*80)

# Process both folders
db_results = process_folder(DB_FOLDER)
db_granular_results = process_folder(DB_GRANULAR_FOLDER)

# Summary
total_files = len(db_results) + len(db_granular_results)
total_keys = sum(r['duplicate_keys'] for r in db_results + db_granular_results)

print(f"\n{'='*80}")
print("FINAL SUMMARY")
print(f"{'='*80}")
print(f"Files with duplicate keys (db): {len(db_results)}")
print(f"Files with duplicate keys (db_granular): {len(db_granular_results)}")
print(f"Total files fixed: {total_files}")
print(f"Total duplicate keys merged: {total_keys}")
print(f"{'='*80}")

if total_files > 0:
    print("\n✓ All files have been fixed and saved!")
else:
    print("\n✓ No duplicate keys found in any files!")
