import pandas as pd
import numpy as np
import gc

# =================================================================
# CONFIGURATION
# =================================================================
FILE_1_BASE = r'base_file.xlsx'
FILE_2_NEW  = r'file_with_extras.xlsx'
OUTPUT_FILE = 'Position_Based_Comparison.xlsx'

# The column you want to validate for 10-12 characters
CHECK_COL   = 'Account_Code' 
FILTERS     = {'Category': ['Hardware']} 
# =================================================================

def run_comparison():
    try:
        print("🚀 Loading files...")
        df1 = pd.read_excel(FILE_1_BASE)
        df2 = pd.read_excel(FILE_2_NEW)

        # 1. Apply Filters
        for col, values in FILTERS.items():
            if col in df1.columns: df1 = df1[df1[col].isin(values)]
            if col in df2.columns: df2 = df2[df2[col].isin(values)]

        # 2. Create a Synthetic ID based on the physical Row Number
        # This forces the merge to compare Row 1 to Row 1, Row 2 to Row 2
        df1['Row_ID'] = range(len(df1))
        df2['Row_ID'] = range(len(df2))
        
        # Capture Excel-style row numbers for the report
        df1['Base_Excel_Row'] = df1.index + 2
        df2['New_Excel_Row'] = df2.index + 2

        # 3. Validation: Length Check on the specific column
        def check_len(val):
            s_val = str(val).strip()
            if s_val == 'nan': return True # Ignore empty cells or handle as you wish
            return 10 <= len(s_val) <= 12

        print("🔍 Checking row-by-row differences...")
        # Merge on our synthetic Row_ID
        df_all = pd.merge(df1, df2, on='Row_ID', how='outer', indicator='Presence', suffixes=('_base', '_new'))

        def analyze_row(row):
            if row['Presence'] == 'left_only': return 'DELETED'
            if row['Presence'] == 'right_only': return 'NEW'
            
            # Check Length Validation for both base and new
            len_ok_base = check_len(row[f'{CHECK_COL}_base'])
            len_ok_new  = check_len(row[f'{CHECK_COL}_new'])
            
            if not len_ok_base or not len_ok_new:
                return f'INVALID LENGTH: {CHECK_COL} must be 10-12 chars'

            # Compare all original columns
            cols_to_compare = [c for c in df1.columns if c not in ['Row_ID', 'Base_Excel_Row']]
            changes = []
            
            for col in cols_to_compare:
                b_val, n_val = row[f'{col}_base'], row[f'{col}_new']
                if str(b_val) != str(n_val) and not (pd.isna(b_val) and pd.isna(n_val)):
                    changes.append(col)
            
            if changes:
                return f"MODIFIED: Changes in {', '.join(changes)}"
            return "COMMON: No Change"

        df_all['Comparison_Result'] = df_all.apply(analyze_row, axis=1)

        # 4. Clean up and Save
        # Reorganize columns to put Result and Row IDs at the front
        cols = ['Comparison_Result', 'Base_Excel_Row', 'New_Excel_Row'] + [c for c in df_all.columns if c not in ['Comparison_Result', 'Base_Excel_Row', 'New_Excel_Row', 'Presence', 'Row_ID']]
        df_final = df_all[cols]

        print(f"💾 Saving to {OUTPUT_FILE}...")
        writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
        df_final.to_excel(writer, index=False)
        
        # (Standard Formatting logic here as used in previous steps...)
        
        writer.close()
        print("✅ DONE!")

    except Exception as e:
        print(f"❌ Error: {e}")

run_comparison()
