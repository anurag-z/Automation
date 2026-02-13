import pandas as pd
import numpy as np

# =================================================================
# CONFIGURATION
# =================================================================
FILE_1_BASE = r'base_file.xlsx'
FILE_2_NEW  = r'file_with_extras.xlsx'
OUTPUT_FILE = 'Final_Single_View_Report.xlsx'

CHECK_COL   = 'Account_Code' # The column to validate (10-12 chars)
FILTERS     = {'Category': ['Hardware']} 
# =================================================================

def run_comparison():
    try:
        print("🚀 Loading and filtering data...")
        df1 = pd.read_excel(FILE_1_BASE)
        df2 = pd.read_excel(FILE_2_NEW)

        # Apply Filters
        for col, values in FILTERS.items():
            if col in df1.columns: df1 = df1[df1[col].isin(values)].reset_index(drop=True)
            if col in df2.columns: df2 = df2[df2[col].isin(values)].reset_index(drop=True)

        # Ensure both dataframes are the same length for row-by-row comparison
        max_rows = max(len(df1), len(df2))
        
        results = []

        print("🔍 Checking for changes and validation errors...")
        for i in range(max_rows):
            # Get rows (handle cases where one file might be shorter)
            row_base = df1.iloc[i] if i < len(df1) else None
            row_new  = df2.iloc[i] if i < len(df2) else None
            
            status = "COMMON"
            change_details = ""
            
            if row_base is not None and row_new is not None:
                # 1. Validation Check (10-12 characters)
                val = str(row_new[CHECK_COL]).strip()
                if not (10 <= len(val) <= 12):
                    status = "INVALID LENGTH"
                
                # 2. Check for changes in any column
                changes = []
                for col in df2.columns:
                    b_val = str(row_base[col]) if col in df1.columns else "N/A"
                    n_val = str(row_new[col])
                    if b_val != n_val:
                        changes.append(col)
                
                if changes:
                    # If it wasn't already marked INVALID, mark it MODIFIED
                    status = "MODIFIED" if status == "COMMON" else status
                    change_details = f"Changed: {', '.join(changes)}"
            
            elif row_base is None:
                status = "NEW ROW"
            else:
                status = "DELETED ROW"

            # Create the final row based on the NEW data, plus our status columns
            final_row = row_new.to_dict() if row_new is not None else row_base.to_dict()
            final_row['Comparison_Status'] = status
            final_row['Change_Logs'] = change_details
            final_row['Original_Row'] = i + 2
            results.append(final_row)

        # 3. Save to Excel
        df_final = pd.DataFrame(results)
        
        # Move status columns to the front
        cols = ['Comparison_Status', 'Change_Logs', 'Original_Row'] + [c for c in df2.columns]
        df_final = df_final[cols]

        print(f"💾 Saving report to {OUTPUT_FILE}...")
        writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
        df_final.to_excel(writer, index=False)
        
        # Formatting
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        red_fmt = workbook.add_format({'bg_color': '#FFC7CE'})   # Errors/Deleted
        yel_fmt = workbook.add_format({'bg_color': '#FFEB9C'})   # Modified
        
        worksheet.conditional_format(1, 0, len(df_final), 0, {
            'type': 'cell', 'criteria': 'containing', 'value': 'INVALID', 'format': red_fmt
        })
        worksheet.conditional_format(1, 0, len(df_final), 0, {
            'type': 'cell', 'criteria': 'containing', 'value': 'MODIFIED', 'format': yel_fmt
        })

        writer.close()
        print("✅ Process Complete!")

    except Exception as e:
        print(f"❌ Error: {e}")

run_comparison()
