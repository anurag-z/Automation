import pandas as pd
import gc
import os

# =================================================================
# CONFIGURATION - Update these paths
# =================================================================
FILE_1_BASE = r'C:\Path\To\base_file.csv'        # Your reference file
FILE_2_NEW  = r'C:\Path\To\file_with_extras.csv' # The file with extras
OUTPUT_FILE = 'Difference_Report.xlsx'
# =================================================================

def run_comparison():
    try:
        # 1. Load the data
        print("🚀 Loading CSV files (this may take a moment for 100MB+)...")
        df1 = pd.read_csv(FILE_1_BASE, low_memory=False)
        df2 = pd.read_csv(FILE_2_NEW, low_memory=False)

        print("🔄 Comparing all columns to find 'In-Between' differences...")
        # Indicator=True creates a column named '_merge'
        # how='outer' ensures we find everything in both files
        df_diff = pd.merge(df1, df2, how='outer', indicator='Comparison_Result')

        # 2. Identify which file has what
        # We rename the labels to be human-readable
        df_diff['Comparison_Result'] = df_diff['Comparison_Result'].map({
            'left_only': 'MISSING: In Base but gone in New',
            'right_only': 'EXTRA: Found in New file (In-Between)',
            'both': 'MATCH'
        })

        # 3. Filter: Keep ONLY the differences (makes the file much smaller)
        df_final = df_diff[df_diff['Comparison_Result'] != 'MATCH'].copy()

        # Clean up memory immediately
        del df1
        del df2
        del df_diff
        gc.collect()

        if df_final.empty:
            print("✅ No differences found! The files are identical.")
            return

        print(f"💾 Found {len(df_final)} differences. Saving to Excel with colors...")
        
        # 4. Write to Excel with Highlighting
        writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
        df_final.to_excel(writer, index=False, sheet_name='Diffs')

        workbook  = writer.book
        worksheet = writer.sheets['Diffs']

        # Define Formats
        red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Missing
        yel_fmt = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}) # Extra

        # Apply formatting to the result column (the last column)
        last_col_idx = len(df_final.columns) - 1
        
        # Highlight MISSING rows Red
        worksheet.conditional_format(1, 0, len(df_final), last_col_idx, {
            'type': 'cell', 'criteria': 'containing', 'value': 'MISSING', 'format': red_fmt
        })
        # Highlight EXTRA rows Yellow
        worksheet.conditional_format(1, 0, len(df_final), last_col_idx, {
            'type': 'cell', 'criteria': 'containing', 'value': 'EXTRA', 'format': yel_fmt
        })

        writer.close()
        print(f"\n✅ SUCCESS! Report generated: {OUTPUT_FILE}")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    run_comparison()
