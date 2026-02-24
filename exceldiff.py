import pandas as pd

# --- Configuration Variables from C# ---
FILE_1_BASE = r"C:\Test\FieldAnalysis_Combined_1040_18_Feb.xlsx"
FILE_2_NEW = r"C:\Test\FieldAnalysis_Combined_1040_test.xlsx"
OUTPUT_FILE = r"C:\Test\difference_report_Full_Compare.xlsx"

FILTERS = {
    "Safe to Extend": ["Yes"],
    "Action Required": ["Move and Extend (Non-Group)", "Extend Only (Non-Group)"]
}

CHECK_COL = "Length"

def run_comparison():
    try:
        print("Loading and filtering data...")
        # 1. Read Excel files
        df1 = pd.read_excel(FILE_1_BASE)
        df2 = pd.read_excel(FILE_2_NEW)

        # 2. Apply Filters to df1 (Base)
        for key, values in FILTERS.items():
            if key in df1.columns:
                # Keep rows where the column value is in our list of allowed values
                df1 = df1[df1[key].astype(str).isin([str(v) for v in values])]

        # 3. Define match columns
        match_cols = ["Area", "Screen Name", "Screen Number", "Field Name", "Level", "Row", "Column"]
        
        # Ensure match columns are strings to avoid type mismatch during merge
        for col in match_cols:
            if col in df1.columns and col in df2.columns:
                df1[col] = df1[col].astype(str).fillna("")
                df2[col] = df2[col].astype(str).fillna("")
            else:
                print(f"Warning: Match column '{col}' is missing from one of the files.")

        # 4. Perform a Left Join to find matches
        df2_with_idx = df2.copy()
        df2_with_idx['Original row'] = df2_with_idx.index + 2 # +2 for header and 0-index offset
        
        merged = pd.merge(
            df1, 
            df2_with_idx, 
            on=match_cols, 
            how='left', 
            suffixes=('_base', '') 
        )

        def process_row(row):
            status = "COMMON"
            change_logs = ""
            
            # Check for DELETED (No match found in df2)
            if pd.isna(row.get('Original row')):
                return pd.Series(["DELETED ROW", "", ""])

            # Check for MODIFIED (Compare all columns present in df2)
            changes = []
            for col in df2.columns:
                base_val = str(row.get(f"{col}_base", "")).strip()
                new_val = str(row.get(col, "")).strip()
                
                # We only flag a change if the base value actually existed and is different
                if pd.notna(row.get(f"{col}_base")) and base_val != new_val:
                    changes.append(col)
            
            if changes:
                status = "MODIFIED"
                change_logs = f"Changed: {', '.join(changes)}"
            
            # Validation Check
            if str(row.get(CHECK_COL, "")).strip() != "10":
                status = "INVALID LENGTH"
                
            return pd.Series([status, change_logs, row['Original row']])

        # 5. Apply the logic across the dataframe
        print("Comparing rows...")
        merged[['Comparison_Status', 'Change_Logs', 'Original row']] = merged.apply(process_row, axis=1)

        # 6. Reorder columns to match your C# output structure
        final_cols = ['Original row', 'Comparison_Status', 'Change_Logs'] + list(df2.columns)
        # Ensure we only select columns that actually exist to prevent KeyError
        final_cols = [col for col in final_cols if col in merged.columns]
        final_df = merged[final_cols]

        # 7. Save to Excel
        print(f"Saving report to {OUTPUT_FILE}...")
        final_df.to_excel(OUTPUT_FILE, index=False)
        print("Process Complete!")

    except Exception as e:
        print(f"X Error: {str(e)}")

# Execute the script
if __name__ == "__main__":
    run_comparison()
