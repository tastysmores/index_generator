import os
import shutil
import pandas as pd
from pathlib import Path
import re


def strip_invalid_characters(name: str, max_length: int = 200) -> str:
    # Remove invalid Windows characters
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    
    # Remove control characters (newlines, tabs, etc.)
    name = re.sub(r'[\x00-\x1f]', '', name)
    
    # Remove trailing spaces and dots
    name = name.rstrip(' .')
    
    return name[:max_length]

def get_unique_path(path):
    counter = 1
    while path.exists():
        path = path.with_name(f"{path.stem} ({counter}){path.suffix}")
        counter += 1
    return path

def copy_renamed_files(excel_file, source_root, destination_root):
    # Load the Excel file
    df = pd.read_excel(excel_file)
    
    # Make sure the destination root exists
    #destination_root = Path(destination_root).resolve()
    #source_root = Path(source_root).resolve()

    # Copy relevant files and their subfolders
    for index, row in df.iterrows():
        
        #original_path = Path(row['Relative Path']).resolve()
        original_path = source_root / Path(row['Relative Path'])
        original_path_parent = Path(row['Relative Path']).parent
        
        new_name = strip_invalid_characters(row['New File Name'])

        destination_path = destination_root / original_path_parent / new_name

        destination_path = get_unique_path(destination_path)
        #print(original_path)
        #print(new_path)
                

        try:
            #relative_path = str(original_path).replace(str(source_root), '').lstrip("\\/")
            #destination_path = destination_root / new_path
            
            #print(destination_path)

            # Copy the .eml file
            destination_path.parent.mkdir(parents=True, exist_ok=True)
            
            shutil.copy2(original_path, destination_path)

            # Check for and copy the associated subfolder
            subfolder = original_path.with_suffix('')
            if subfolder.is_dir():
                subfolder_destination = destination_root / subfolder.relative_to(source_root)
                shutil.copytree(subfolder, subfolder_destination, dirs_exist_ok=True)

            print(f"Copied: {original_path} -> {destination_path}")

        except Exception as e:
            print(f"Error copying '{original_path}': {e}")

    print("Copy completed.")

#if __name__ == "__main__":
#    excel_file = input("Enter the path to the Excel file: ").strip()
#    source_root = input("Enter the path to the source root folder: ").strip()
#    destination_root = input("Enter the path to the destination root folder: ").strip()

#    copy_relevant_files(excel_file, source_root, destination_root)
