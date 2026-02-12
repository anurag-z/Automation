import pandas as pd
import gc

# =================================================================
# CONFIGURATION
# =================================================================
FILE_1_BASE = r'C:\Program Files\base_file.csv'        
FILE_2_NEW  = r'C:\Path\To\file_with_extras.csv' 
OUTPUT_FILE = 'Row_Shift_Analysis.xlsx'
# =================================================================

def run_shift_analysis():
    try:
        print("🚀 Loading files...")
        df1 = pd.read_csv(FILE_1_BASE, low_memory=False)
        df2 = pd.read_csv(FILE_2_NEW, low_memory=False)

        # 1. Capture original row numbers (Excel style)
        df1['Base_Row_#'] = df1.index + 2
        df2['New_Row_#'] = df2.index + 2

        print("🔍 Searching entire files for matches (Position-Independent)...")
        # Columns to match (everything except our helper row number columns)
        cols = [c for c in df1.columns if c not in ['Base_Row_#', 'New_Row_#']]
        
        # Merge finds where the DATA matches, regardless of where it is sitting
        df_all = pd.merge(df1, df2, on=cols, how='outer', indicator='Presence')

        # 2. Define Logic for "Shifted" vs "Static" vs "Missing"
        conditions = [
            (df_all['Presence'] == 'both') & (df_all['Base_Row_#'] == df_all['New_Row_#']),
            (df_all['Presence'] == 'both') & (df_all['Base_Row_#'] != df_all['New_Row_#']),
            (df_all['Presence'] == 'left_only'),
            (df_all['Presence'] == 'right_only')
        ]
        choices = [
            'STAYED: Same data, Same row',
            'SHIFTED: Same data, Different row',
            'DELETED: Data found in Base only',
            'NEW: Data found in New file only'
        ]
        
        df_all['Result_Status'] = pd.np.select(conditions, choices, default='Unknown')

        # 3. Filter: We only want to see things that changed or moved
        df_final = df_all[df_all['Result_Status'] != 'STAYED: Same data, Same row'].copy()

        # Sort so shifted rows appear logically
        df_final = df_final.sort_values(by=['Result_Status', 'New_Row_#'])

        del df1, df2, df_all
        gc.collect()

        print(f"💾 Saving {len(df_final)} changes to Excel...")
        writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
        df_final.to_excel(writer, index=False, sheet_name='Shift_Analysis')

        workbook  = writer.book
        worksheet = writer.sheets['Shift_Analysis']

        # Formats
        blue_fmt = workbook.add_format({'bg_color': '#DDEBF7', 'font_color': '#003366'}) # Shifted
        red_fmt  = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Deleted
        yel_fmt  = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}) # New

        last_col = len(df_final.columns) - 1
        
        # Apply Highlighting based on Status
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
        print(f"✅ DONE! Analyze shifted rows here: {OUTPUT_FILE}")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    run_shift_analysis()
