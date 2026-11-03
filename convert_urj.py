import pandas as pd
import sys
import os
import re
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Font, PatternFill, Side

def format_as_table(output_path, df, original_filename):
    """Helper function to apply Table formatting, Auto-fit widths, and Save"""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        
        # --- SHEET NAME LOGIC ---
        sheet_name = re.sub(r'[\\/*?:\[\]]', '_', original_filename)[:31]
        
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        max_row, max_col = len(df) + 1, len(df.columns)
        ref = f"A1:{get_column_letter(max_col)}{max_row}"

        clean_tab_name = re.sub(r'[^a-zA-Z0-9]', '_', original_filename)[:30]
        if not clean_tab_name or not clean_tab_name[0].isalpha():
            clean_tab_name = "T_" + clean_tab_name
            
        tab = Table(displayName=clean_tab_name, ref=ref)
        
        style = TableStyleInfo(
            name="TableStyleMedium9", 
            showFirstColumn=False,
            showLastColumn=False, 
            showRowStripes=True, 
            showColumnStripes=False
        )
        tab.tableStyleInfo = style
        worksheet.add_table(tab)

        header_fill = PatternFill(fill_type="solid", fgColor="008000")
        for cell in worksheet[1][:max_col]:
            cell.fill = header_fill
            cell.font = Font(name=cell.font.name, size=cell.font.sz, bold=True, color="FFFFFF")

        green_fill = PatternFill(fill_type="solid", fgColor="D9EFD3")
        white_fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
        green_bottom = Side(style="thin", color="70E050")
        gray_right = Side(style="thin", color="D9D9D9")
        for row_idx in range(2, max_row + 1):
            fill = green_fill if row_idx % 2 == 0 else white_fill
            for cell in worksheet[row_idx][:max_col]:
                cell.fill = fill
                cell.border = Border(bottom=green_bottom, right=gray_right)

        # --- SAFE AUTO-FIT COLUMN WIDTHS ---
        # Iterate over worksheet columns directly to avoid issues with duplicate/NaN headers
        for col_idx, col_cells in enumerate(worksheet.iter_cols(min_row=1, max_row=max_row), start=1):
            max_len = 0
            for cell in col_cells:
                if cell.value is not None:
                    # Measure the string length of each cell
                    cell_len = len(str(cell.value))
                    if cell_len > max_len:
                        max_len = cell_len
            
            # +5 to +6 padding ensures the Excel table filter dropdown arrow doesn't cover text
            # Min width of 12 ensures small headers aren't squeezed
            col_letter = get_column_letter(col_idx)
            worksheet.column_dimensions[col_letter].width = max(max_len + 5, 12)

def process_item(item_path):
    """Handles individual file conversion while keeping name identical"""
    if item_path.lower().endswith('.urj'):
        try:
            base_name = os.path.splitext(item_path)[0]
            file_name_only = os.path.basename(base_name)
            
            print(f"Processing: {file_name_only}.urj")
            
            # --- Robust encoding logic ---
            # dtype=str prevents long account numbers from turning into scientific notation (e.g. 2E+12)
            try:
                df = pd.read_csv(item_path, sep='~', encoding='utf-8', on_bad_lines='skip', dtype=str)
            except UnicodeDecodeError:
                print(f"   - Note: UTF-8 decoding failed. Trying Windows-1252 encoding...")
                df = pd.read_csv(item_path, sep='~', encoding='cp1252', on_bad_lines='skip', dtype=str)
            
            # Ensure all column headers are clean strings (replaces NaN with 'Unnamed')
            df.columns = [str(c) if pd.notna(c) and str(c).strip() != '' else f"Column_{i+1}" for i, c in enumerate(df.columns)]

            output_path = base_name + ".xlsx"
            format_as_table(output_path, df, file_name_only)
            
            print(f"   - Success: {os.path.basename(output_path)}")
        except Exception as e:
            print(f"   - Error processing {item_path}: {e}")
    else:
        print(f"   - Warning: {os.path.basename(item_path)} is not a .urj file and will be skipped.")

def main():
    if len(sys.argv) < 2:
        print("Error: No files or folders detected.")
        return

    for input_item in sys.argv[1:]:
        if os.path.isdir(input_item):
            print(f"Reading Folder: {input_item}")
            for root, dirs, files in os.walk(input_item):
                for file in files:
                    if file.lower().endswith('.urj'):
                        process_item(os.path.join(root, file))
        else:
            process_item(input_item)

    print("\n--- All Tasks Complete! ---")

if __name__ == "__main__":
    main()