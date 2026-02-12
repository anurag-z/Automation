import pandas as pd
import numpy as np 
import gc

# =================================================================
# CONFIGURATION
# =================================================================
FILE_1_BASE = r'C:\Program Files\base_file.csv'        
FILE_2_NEW  = r'C:\Path\To\file_with_extras.csv' 
OUTPUT_FILE = 'Full_Data_Comparison_Report.xlsx'
# =================================================================

def run_full_comparison():
    try:
        print("🚀 Loading files...")
        df1 = pd.read_csv(FILE_1_BASE, low_memory=False)
        df2 = pd.read_csv(FILE_2_NEW, low_memory=False)

        # 1. Capture original row numbers (Excel style: Index + 2)
        df1['Base_Row_#'] = df1.index + 2
        df2['New_Row_#'] = df2.index + 2

        print("🔍 Comparing data across all rows...")
        # Get all column names except our helper row numbers
        cols = [c for c in df1.columns if c not in ['Base_Row_#', 'New_Row_#']]
        
        # Merge finds exact matches anywhere in the file
        df_all = pd.merge(df1, df2, on=cols, how='outer', indicator='Presence')

        # 2. Logic to categorize every row
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

        # --- NO FILTERING ---
        # We keep all rows (COMMON, SHIFTED, DELETED, NEW)
        df_final = df_all.sort_values(by=['New_Row_#', 'Base_Row_#'])

        del df1, df2, df_all
        gc.collect()

        print(f"💾 Saving {len(df_final)} rows to Excel...")
        writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
        df_final.to_excel(writer, index=False, sheet_name='Full_Comparison')

        workbook  = writer.book
        worksheet = writer.sheets['Full_Comparison']

        # Formatting Colors
        green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}) # Common
        blue_fmt  = workbook.add_format({'bg_color': '#DDEBF7', 'font_color': '#003366'}) # Shifted
        red_fmt   = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Deleted
        yel_fmt   = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}) # New

        last_col = len(df_final.columns) - 1
        
        # Apply Highlighting
        worksheet.conditional_format(1, 0, len(df_final), last_col, {
            'type': 'cell', 'criteria': 'containing', 'value': 'COMMON', 'format': green_fmt
        })
        worksheet.conditional_format(1, 0, len(df_final), last_col, {
            'type': 'cell', 'criteria': 'containing', 'value': 'SHIFTED', 'format': blue_fmt
        })
        worksheet.conditional_format(1, 0, len(df_final), last_col, {
            'type': 'cell', 'criteria': 'containing', 'value': 'DELETED', 'format': red_fmt
        })
        worksheet.conditional_format(1, 0, len(df_final), last_col, {
            'type': 'cell', 'criteria': 'containing', 'value': 'NEW', 'format': yel_fmt
        })

        writer.close()
        print(f"✅ DONE! Full report saved: {OUTPUT_FILE}")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    run_full_comparison()
