import re
import json
import pandas as pd

def lua_to_json(lua_str):
    """
    Converts a Lua table string to a JSON string.
    """
    content = lua_str.strip()

    # 1. Remove the 'return' statement
    if content.startswith("return {"):
        content = content[7:]
    
    # 2. Remove Lua comments (single line -- and block --[[ ... ]])
    content = re.sub(r'--\[\[.*?\]\]', '', content, flags=re.DOTALL)
    content = re.sub(r'--.*', '', content)

    # 3. Replace Lua-specific values with JSON equivalents
    content = content.replace("nil", "null")

    # 4. Fix Keys: Convert ["Key"] = to "Key":
    content = re.sub(r'\["(.*?)"\]\s*=', r'"\1":', content)

    # 5. Fix Trailing Commas
    # Remove commas that appear before a closing brace } or bracket ]
    content = re.sub(r',(\s*[}\]])', r'\1', content)

    # 6. Clean up whitespace
    content = re.sub(r'\s+', ' ', content).strip()

    return content

def extract_hero_info(data):
    extracted_list = []
    
    for hero_name, hero_data in data.items():
        if not isinstance(hero_data, dict):
            continue
            
        # 1. Basic Info
        hero_id = hero_data.get("id", "")
        name = hero_data.get("name", hero_name)
        
        # Handle Roles
        role1 = hero_data.get("role1", "")
        role2 = hero_data.get("role2", "")
        role = f"{role1} / {role2}" if role2 else role1
        
        # Handle Lanes
        lane1 = hero_data.get("lane1", "")
        lane2 = hero_data.get("lane2", "")
        lane = f"{lane1} / {lane2}" if lane2 else lane1
        
        # Handle Specialties
        spec1 = hero_data.get("specialty1", "")
        spec2 = hero_data.get("specialty2", "")
        speciality = f"{spec1} / {spec2}" if spec2 else spec1
        
        # 2. Maximum Stats (Level 15)
        stats_block = hero_data.get("stats", {})
        if not isinstance(stats_block, dict):
            stats_block = {}
            
        max_stats = {
            "hp_max": stats_block.get("hp15", ""),
            "hp_regen_max": stats_block.get("hp_regen15", ""),
            "mana_max": stats_block.get("mana15", ""),
            "mana_regen_max": stats_block.get("mana_regen15", ""),
            "physical_atk_max": stats_block.get("physical_atk15", ""),
            "physical_def_max": stats_block.get("physical_def15", ""),
            "magic_def_max": stats_block.get("magic_def15", ""),
            "atk_spd_max": stats_block.get("atk_spd15", ""),
            "movement_spd": stats_block.get("movement_spd", ""),
            # ADDED: Basic Attack Range
            "basic_atk_range": stats_block.get("basic_atk_range", "")
        }

        # 3. Construct the flat record
        entry = {
            "id": hero_id,
            "name": name,
            "role": role,
            "lane": lane,
            "speciality": speciality,
            **max_stats 
        }
        
        extracted_list.append(entry)
        
    return extracted_list

def main():
    input_file = 'Hero.txt'
    output_json = 'extracted_heroes.json'
    output_excel = 'extracted_heroes.xlsx'

    print(f"Reading data from {input_file}...")
    
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            lua_data = f.read()
    except FileNotFoundError:
        print(f"Error: The file '{input_file}' was not found.")
        return

    print("Converting Lua to JSON format...")
    json_str = lua_to_json(lua_data)
    
    try:
        hero_dict = json.loads(json_str)
    except json.JSONDecodeError as e:
        print(f"Error parsing converted JSON: {e}")
        return

    if not hero_dict:
        print("Error: Parsed dictionary is empty.")
        return

    print("Extracting hero information...")
    heroes = extract_hero_info(hero_dict)
    
    print(f"Successfully extracted {len(heroes)} heroes.")

    # Save to JSON
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(heroes, f, indent=4, ensure_ascii=False)
    print(f"Saved JSON to {output_json}")

    # Save to Excel
    try:
        if not heroes:
            print("No heroes to save to Excel.")
            return

        df = pd.DataFrame(heroes)
        
        # Define columns including the new basic_atk_range
        cols = ["id", "name", "role", "lane", "speciality", 
                "hp_max", "hp_regen_max", "mana_max", "mana_regen_max", 
                "physical_atk_max", "physical_def_max", "magic_def_max", 
                "atk_spd_max", "movement_spd", "basic_atk_range"]
        
        existing_cols = [c for c in cols if c in df.columns]
        df = df[existing_cols]
        df.to_excel(output_excel, index=False)
        print(f"Saved Excel to {output_excel}")
    except Exception as e:
        print(f"Error saving Excel: {e}")

if __name__ == "__main__":
    main()