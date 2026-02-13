import pandas as pd
import numpy as np
import gc

# =================================================================
# CONFIGURATION
# =================================================================
FILE_1_BASE = r'C:\Path\To\base_file.xlsx'
FILE_2_NEW  = r'C:\Path\To\file_with_extras.xlsx'
OUTPUT_FILE = 'Filtered_Data_Comparison_Report.xlsx'

# Define your filters here (Column Name: [List of values to keep])
# Example: {'Status': ['Active', 'Pending'], 'Region': ['North']}
FILTERS = {
    'Status': ['Active'],
    'Department': ['Sales', 'IT']
}
# =================================================================

def run_full_comparison():
    try:
        print("🚀 Loading Excel files...")
        # Use openpyxl engine for modern .xlsx files
        df1 = pd.read_excel(FILE_1_BASE, engine='openpyxl')
        df2 = pd.read_excel(FILE_2_NEW, engine='openpyxl')

        # 1. Apply Filters
        print("✂️ Applying filters to both datasets...")
        for col, values in FILTERS.items():
            if col in df1.columns:
                df1 = df1[df1[col].isin(values)]
            if col in df2.columns:
                df2 = df2[df2[col].isin(values)]

        # 2. Capture original row numbers (Excel style: Index + 2)
        # We do this AFTER filtering so we know the physical row in the source file
        df1['Base_Row_#'] = df1.index + 2
        df2['New_Row_#'] = df2.index + 2

        print("🔍 Comparing data across all rows...")
        cols = [c for c in df1.columns if c not in ['Base_Row_#', 'New_Row_#']]
        
        # Merge logic (Inner/Outer check)
        df_all = pd.merge(df1, df2, on=cols, how='outer', indicator='Presence')

        # 3. Categorize
        conditions = [
            (df_all['Presence'] == 'both') & (df_all['Base_Row_#'] == df_all['New_Row_#']),
            (df_all['Presence'] == 'both') & (df_all['Base_Row_#'] != df_all['New_Row_#']),
            (df_all['Presence'] == 'left_only'),
            (df_all['Presence'] == 'right_only')
        ]
        choices = [
            'COMMON: Same data, Same row',
            'SHIFTED: Same data, Moved row',
            'DELETED: Data in Base only',
            'NEW: Data in New file only'
        ]
        
        df_all['Comparison_Result'] = np.select(conditions, choices, default='Unknown')
        df_final = df_all.sort_values(by=['New_Row_#', 'Base_Row_#'])

        del df1, df2, df_all
        gc.collect()

        print(f"💾 Saving {len(df_final)} filtered rows to Excel...")
        writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
        df_final.to_excel(writer, index=False, sheet_name='Full_Comparison')

        workbook  = writer.book
        worksheet = writer.sheets['Full_Comparison']

        # Formatting
        green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        blue_fmt  = workbook.add_format({'bg_color': '#DDEBF7', 'font_color': '#003366'})
        red_fmt   = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        yel_fmt   = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})

        last_col = len(df_final.columns) - 1
        
        # Apply Highlighting & Add Filter to Header
        worksheet.autofilter(0, 0, len(df_final), last_col)
        worksheet.freeze_panes(1, 0)
        
        for val, fmt in zip(['COMMON', 'SHIFTED', 'DELETED', 'NEW'], [green_fmt, blue_fmt, red_fmt, yel_fmt]):
            worksheet.conditional_format(1, 0, len(df_final), last_col, {
                'type': 'cell', 'criteria': 'containing', 'value': val, 'format': fmt
            })

        writer.close()
        print(f"✅ DONE! Filtered report saved: {OUTPUT_FILE}")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    run_full_comparison()
