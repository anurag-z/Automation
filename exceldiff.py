import pandas as pd
import gc

# =================================================================
# CONFIGURATION
# =================================================================
FILE_1_BASE = r'C:\Program Files\base_file.csv'        
FILE_2_NEW  = r'C:\Path\To\file_with_extras.csv' 
OUTPUT_FILE = 'Difference_Report_With_Row_Nums.xlsx'
# =================================================================

def run_comparison():
    try:
        print("🚀 Loading files...")
        df1 = pd.read_csv(FILE_1_BASE, low_memory=False)
        df2 = pd.read_csv(FILE_2_NEW, low_memory=False)

        # Create a 'Row_Number' column for both files before merging
        # We add +2 because Excel starts at 1 and has a header row
        df1['Base_Row_#'] = df1.index + 2
        df2['New_Row_#'] = df2.index + 2

        print("🔄 Analyzing differences and row positions...")
        # Merge using all columns except the new row numbers
        cols_to_match = [c for c in df1.columns if c not in ['Base_Row_#', 'New_Row_#']]
        
        df_diff = pd.merge(df1, df2, on=cols_to_match, how='outer', indicator='Origin')

        # Create easy-to-read labels
        df_diff['Origin'] = df_diff['Origin'].map({
            'left_only': 'MISSING: In Base but gone in New',
            'right_only': 'EXTRA: Found in New file (In-Between)',
            'both': 'MATCH'
        })

        # Filter out exact matches
        df_final = df_diff[df_diff['Origin'] != 'MATCH'].copy()

        # Clean up memory
        del df1, df2, df_diff
        gc.collect()

        if df_final.empty:
            print("✅ No differences found!")
            return

        print(f"💾 Saving {len(df_final)} differences with row numbers...")
        
        writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
        df_final.to_excel(writer, index=False, sheet_name='Differences')

        workbook  = writer.book
        worksheet = writer.sheets['Differences']

        # Fix for the .add_format Pylance warning (it works regardless of the red line)
        red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        yel_fmt = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})

        last_col_idx = len(df_final.columns) - 1
        
        worksheet.conditional_format(1, 0, len(df_final), last_col_idx, {
            'type': 'cell', 'criteria': 'containing', 'value': 'MISSING', 'format': red_fmt
        })
        worksheet.conditional_format(1, 0, len(df_final), last_col_idx, {
            'type': 'cell', 'criteria': 'containing', 'value': 'EXTRA', 'format': yel_fmt
        })

        writer.close()
        print(f"✅ DONE! Report saved: {OUTPUT_FILE}")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    run_comparison()
