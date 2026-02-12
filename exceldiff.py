import pandas as pd
import gc

# =================================================================
# CONFIGURATION
# =================================================================
FILE_1_BASE = r'C:\Program Files\base_file.csv'        
FILE_2_NEW  = r'C:\Path\To\file_with_extras.csv' 
OUTPUT_FILE = 'Audit_Change_Report.xlsx'
# =================================================================

def run_audit():
    try:
        print("🚀 Loading files...")
        df1 = pd.read_csv(FILE_1_BASE, low_memory=False)
        df2 = pd.read_csv(FILE_2_NEW, low_memory=False)

        # Record original row numbers (Excel style: +2 for header/index)
        df1['Base_Row_#'] = df1.index + 2
        df2['New_Row_#'] = df2.index + 2

        print("🔍 Comparing rows...")
        # Identify columns to compare (exclude the row number helpers)
        cols = [c for c in df1.columns if c not in ['Base_Row_#', 'New_Row_#']]
        
        # Outer join to find all discrepancies
        df_diff = pd.merge(df1, df2, on=cols, how='outer', indicator='Action_Type')

        # Filter out exact matches
        df_final = df_diff[df_diff['Action_Type'] != 'both'].copy()

        # Rename labels for the "Audit" view
        df_final['Action_Type'] = df_final['Action_Type'].map({
            'left_only': 'DELETED or MODIFIED (Old)',
            'right_only': 'ADDED or MODIFIED (New)'
        })

        # --- THE SORTING TRICK ---
        # Sorting by the data columns puts "Old" and "New" versions of a 
        # changed row right next to each other.
        df_final = df_final.sort_values(by=cols)

        # Cleanup
        del df1, df2, df_diff
        gc.collect()

        print(f"💾 Saving Audit Report ({len(df_final)} rows)...")
        writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
        df_final.to_excel(writer, index=False, sheet_name='Audit_Trail')

        workbook  = writer.book
        worksheet = writer.sheets['Audit_Trail']

        # Formatting colors
        red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Missing/Old
        yel_fmt = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}) # Extra/New

        last_col = len(df_final.columns) - 1
        
        # Highlight Red (Old/Deleted)
        worksheet.conditional_format(1, 0, len(df_final), last_col, {
            'type': 'cell', 'criteria': 'containing', 'value': 'DELETED', 'format': red_fmt
        })
        # Highlight Yellow (New/Added)
        worksheet.conditional_format(1, 0, len(df_final), last_col, {
            'type': 'cell', 'criteria': 'containing', 'value': 'ADDED', 'format': yel_fmt
        })

        writer.close()
        print(f"✅ DONE! Audit report ready: {OUTPUT_FILE}")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    run_audit()
