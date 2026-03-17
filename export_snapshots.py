"""
Automated snapshot export from Google Sheets to local files.

This script:
1. Downloads the Excel file from Google Drive using direct download link
2. Updates food_dict.csv from the 'dict' sheet
3. Creates/overwrites daily CSV from the 'food' sheet in macros/ directory

Note: Tanita scale measurements are exported manually to measurements/ directory
"""

import os
from datetime import datetime
import pandas as pd
import requests

# Configuration
PROJECT_DIR = r'c:\git\macros'
MACROS_DIR = os.path.join(PROJECT_DIR, 'macros')
LINK_FILE = os.path.join(PROJECT_DIR, 'link.txt')

# Sheet names
FOOD_SHEET = 'food'
DICT_SHEET = 'dict'


def get_download_link():
    """Read the download link from link.txt."""
    if not os.path.exists(LINK_FILE):
        raise FileNotFoundError(f"link.txt not found at {LINK_FILE}")
    
    with open(LINK_FILE, 'r') as f:
        link = f.read().strip()
    
    if not link:
        raise ValueError("link.txt is empty")
    
    return link


def download_xlsx():
    """Download the Excel file from Google Drive."""
    print("Downloading Excel file from Google Drive...")
    
    link = get_download_link()
    xlsx_path = os.path.join(PROJECT_DIR, 'macros.xlsx')
    
    try:
        response = requests.get(link, timeout=30)
        response.raise_for_status()
        
        with open(xlsx_path, 'wb') as f:
            f.write(response.content)
        
        print(f"✓ Downloaded to {xlsx_path}")
        return xlsx_path
        
    except requests.exceptions.RequestException as e:
        print(f"✗ Download failed: {e}")
        raise


def update_food_dict(xlsx_path):
    """Update food_dict.csv from the 'dict' sheet."""
    print("Updating food_dict.csv...")
    
    try:
        df = pd.read_excel(xlsx_path, sheet_name=DICT_SHEET)
        csv_path = os.path.join(PROJECT_DIR, 'food_dict.csv')
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        
        print(f"✓ Updated food_dict.csv ({len(df)} items)")
        
    except Exception as e:
        print(f"✗ Failed to update food_dict.csv: {e}")
        raise


def create_daily_csv(xlsx_path):
    """Create/overwrite daily CSV with current date in macros/ directory."""
    print("Creating daily macro CSV...")
    
    try:
        # Ensure macros directory exists
        os.makedirs(MACROS_DIR, exist_ok=True)
        
        df = pd.read_excel(xlsx_path, sheet_name=FOOD_SHEET)
        
        # Get current date in format: yyyy-mm-dd
        today = datetime.now()
        date_str = today.strftime('%Y-%m-%d')
        
        csv_path = os.path.join(MACROS_DIR, f'{date_str}.csv')
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        
        print(f"✓ Created/updated macros/{date_str}.csv")
        
    except Exception as e:
        print(f"✗ Failed to create daily CSV: {e}")
        raise


def main():
    """Main execution function."""
    print("=" * 50)
    print("Macro Snapshot Export Tool")
    print("=" * 50)
    print()
    
    try:
        # Download the Excel file
        xlsx_path = download_xlsx()
        print()
        
        # Execute export tasks
        update_food_dict(xlsx_path)
        create_daily_csv(xlsx_path)
        
        print()
        print("=" * 50)
        print("✓ All snapshots exported successfully!")
        print("=" * 50)
        
    except Exception as e:
        print(f"\n✗ Error: {e}")
        raise


if __name__ == '__main__':
    main()
