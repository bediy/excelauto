# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an Excel automation system for managing driving school assessment personnel data (驾校考核人员信息). The system reads personnel information from a source workbook and generates individualized worksheets based on a template, with support for photo insertion.

**Key Data Files:**
- `2025年驾校考核人员信息汇总.xlsx` - Source data containing personnel information
- `工作簿.xlsx` - Target workbook where individual sheets are generated
- `images/` - Directory containing personnel photos (format: `<Name><index>.jpg`)

## Architecture

### Script Workflow

1. **create_person_sheets.py** - Main sheet generation script
   - Reads personnel data from source workbook (Sheet1, starting row 3)
   - Extracts: 姓名 (name, col C), 士兵证号 (soldier ID, col D), 身份证号 (ID card, col E)
   - Uses "肖龙飞" sheet as the template in target workbook
   - Creates or updates individual sheets for each person
   - Populates: B3 (name), D3 (soldier ID digits only), B4 (ID card)

2. **insert_images.py** (referenced but not present in current codebase)
   - Embeds up to 2 photos per person sheet
   - Matches photos by filename prefix against sheet names
   - Sources images from `images/` directory

3. **remove_extra_sheets.py** - Workbook reset utility
   - Keeps only the first worksheet, removes all others
   - Useful for resetting template before regeneration
   - Usage: `python remove_extra_sheets.py <workbook_path>`

### Data Processing Logic

- **extract_digits()** in create_person_sheets.py:120-117 - Strips non-numeric characters from soldier IDs using regex
- Sheet names match person names exactly; sheets are reused if they already exist
- Template worksheet must be named "肖龙飞" or script will raise ValueError
- Empty/None values are safely handled with string conversion and stripping

## Development Commands

### Setup
```bash
python -m venv .venv
.\.venv\Scripts\activate  # Windows
pip install openpyxl Pillow
```

### Running Scripts
```bash
# Generate/refresh person sheets from source data
python create_person_sheets.py

# Insert photos (when insert_images.py exists)
python insert_images.py 工作簿.xlsx images/

# Reset workbook to template-only
python remove_extra_sheets.py 工作簿.xlsx
```

### Testing Workflow
1. Always backup `工作簿.xlsx` before running automation
2. Clone production workbook for testing
3. Verify sheet count and data accuracy after create_person_sheets.py
4. Spot-check random sheets for data fidelity
5. Use Excel's "Inspect Document" to check for warnings

## Code Conventions

- **File paths**: Hardcoded in script constants (INFO_FILE, TARGET_FILE) - update at top of each script for different environments
- **PEP 8 compliance**: 4-space indents, snake_case functions, type hints required
- **Comments**: Use Chinese for domain-specific Excel logic
- **Error handling**: Scripts raise ValueError for missing templates; validate column positions when onboarding new data files
- **Path handling**: Use `pathlib.Path` for all filesystem operations

## Image Asset Requirements

- Format: `<Name><index>.jpg` (e.g., `肖龙飞1.jpg`, `肖龙飞2.jpg`)
- Target size: ~500 KB for optimal performance
- Preferred formats: JPG for photos, PNG for signatures
- Ordering: Numeric suffix determines insertion order (1, 2, etc.)

## Dependencies

- **openpyxl** - Excel file manipulation (.xlsx reading/writing)
- **Pillow** - Image processing for photo insertion
- **Python 3.12+** - Current environment uses Python 3.12.10

## Important Notes

- Scripts modify workbooks in-place; always backup before running
- Column positions in source data are critical: C=name, D=soldier_id, E=id_card (row 3+)
- Template sheet name "肖龙飞" is hardcoded; changing requires code update
- Soldier ID extraction removes all non-digits via regex
