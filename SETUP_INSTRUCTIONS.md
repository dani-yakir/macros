# Macro Snapshot Export - Setup Instructions

This script automatically downloads your Google Sheets macro tracker and exports it to local files.

## Prerequisites

- Python 3.7 or higher
- Direct download link to your Google Sheets file (already in `link.txt`)

## Setup Steps

### 1. Install Dependencies

```bash
cd c:\git\macros
pip install -r requirements.txt
```

That's it! No API setup needed.

### 2. Verify Download Link

The download link is already configured in `link.txt`. If you need to update it:

1. Open your Google Sheets
2. Go to File > Share > Publish to web
3. Choose "Microsoft Excel (.xlsx)" format
4. Get the download link
5. Update `link.txt` with the new link

## Usage

Run anytime to create a snapshot:

```bash
python export_snapshots.py
```

## What It Does

Each time you run the script:

1. **Downloads `macros.xlsx`**: Downloads the latest version from Google Drive
2. **Updates `food_dict.csv`**: Syncs your master food database from the 'dict' sheet
3. **Creates daily CSV**: Named `yyyy-mm-dd.csv` (e.g., `2025-12-26.csv`)
   - Overwrites if run multiple times same day
   - Contains your daily food log from the 'food' sheet

**Tip**: Run this at the end of each day to capture your daily intake.

## Troubleshooting

### "link.txt not found"
- Make sure `link.txt` exists in `c:\git\macros\`
- It should contain your Google Drive download link

### "Download failed"
- Check your internet connection
- Verify the link in `link.txt` is still valid
- Make sure the Google Sheets file is shared properly

### "Failed to read sheet"
- Make sure your Excel file has sheets named 'food' and 'dict'
- Check that the downloaded file isn't corrupted

## Files Created

- `macros.xlsx` - Full workbook download
- `food_dict.csv` - Food database (from 'dict' sheet)
- `yyyy-mm-dd.csv` - Daily food logs (from 'food' sheet, e.g., `2025-12-26.csv`)
