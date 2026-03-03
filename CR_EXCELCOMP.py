import pandas as pd

# --- Configuration ---
# Inputs must be CSV for speed, Output MUST be .xlsx to support cell coloring
FILE_1_BASE = r"C:\Test\FieldAnalysis_Combined_1040_18_Feb.csv"
FILE_2_NEW = r"C:\Test\FieldAnalysis_Combined_1040_test.csv"
OUTPUT_FILE = r"C:\Test\Python_excel\difference_report_Full_Compare.xlsx"

FILTERS = {
    "Safe to Extend": ["Yes"],
    "Action Required": ["Move and Extend (Non-Group)", "Extend Only (Non-Group)"]
}
CHECK_COL = "Length"

# --- NEW: Define the End columns to compare ---
BASE_END_COL = "current end"
TEST_END_COL = "suggest end"

# --- Define the range to compare (0-based indexing) ---
START_RECORD = 0
END_RECORD = 500  # Set to None to process to the very end: END_RECORD = None

def run_comparison():
    try:
        print("Loading CSV data...")
        # FIX: Read as object to prevent int crashes, then convert to string
        df1 = pd.read_csv(FILE_1_BASE, dtype=object).fillna("").astype(str)
        df2 = pd.read_csv(FILE_2_NEW, dtype=object).fillna("").astype(str)

        # --- Capture original row numbers before any filtering ---
        df1['original row_base'] = df1.index + 2

        # 1. Apply Filters to df1
        print("Applying filters...")
        for key, values in FILTERS.items():
            if key in df1.columns:
                df1 = df1[df1[key].isin([str(v) for v in values])]

        # 2. Apply the Range Slice AFTER filtering
        total_filtered = len(df1)
        print(f"Total Base records after filtering: {total_filtered}")
        
        df1 = df1.iloc[START_RECORD:END_RECORD]
        actual_end = min(END_RECORD if END_RECORD else total_filtered, total_filtered)
        print(f"Comparing a range of {len(df1)} records (Index {START_RECORD} to {actual_end})...")

        # 3. Define match columns and prep for merge
        # FIX: Added 'Field Number' to guarantee unique rows
        match_cols = ["Area", "Screen Name", "Screen Number", "Field Name", "Field Number", "Level", "Row", "Column"]
        for col in match_cols:
            if col in df1.columns and col in df2.columns:
                df1[col] = df1[col].str.strip()
                df2[col] = df2[col].str.strip()

        # --- Capture original row numbers for test data ---
        df2_with_idx = df2.copy()
        df2_with_idx['original row_test'] = df2_with_idx.index + 2 
        
        # 4. Perform Left Join
        merged = pd.merge(df1, df2_with_idx, on=match_cols, how='left', suffixes=('_base', ''))

        # 5. Row processing logic
        def process_row(row):
            status = "SUCCESS" 
            change_logs = []
            
            # Check for DELETED 
            if pd.isna(row.get('original row_test')) or str(row.get('original row_test')).strip() == "nan" or row.get('original row_test') == "":
                return pd.Series(["DELETED ROW", ""])

            # --- A. Check END Column Validation ---
            base_end = str(row.get(f"{BASE_END_COL}_base", "")).strip()
            test_end = str(row.get(TEST_END_COL, "")).strip()
            end_mismatch = (base_end != test_end)
            if end_mismatch:
                change_logs.append(f"END MISMATCH (Expected '{base_end}', Got '{test_end}')")

            # --- B. Check LENGTH Validation ---
            length_val = str(row.get(CHECK_COL, "")).strip()
            length_invalid = (length_val != "12")
            if length_invalid:
                change_logs.append(f"INVALID LENGTH ({length_val})")

            # --- C. Check for general MODIFIED columns ---
            changes = []
            for col in df2.columns:
                # Skip tracking/validation columns
                if col in match_cols or col == 'original row_test' or col == CHECK_COL or col == TEST_END_COL:
                    continue
                
                base_val = str(row.get(f"{col}_base", "")).strip()
                new_val = str(row.get(col, "")).strip()
                
                if base_val != new_val:
                    changes.append(col)
            
            if changes:
                change_logs.insert(0, f"Changed: {', '.join(changes)}")

            # --- Determine Final Status ---
            if length_invalid or end_mismatch:
                status = "VALIDATION FAILED"
            elif changes:
                status = "MODIFIED"
            else:
                status = "SUCCESS" # Everything perfectly matches expectations!
                
            final_logs = " | ".join(change_logs)
            return pd.Series([status, final_logs])

        print("Checking for differences...")
        merged[['Comparison_Status', 'Change_Logs']] = merged.apply(process_row, axis=1)

        # 6. Build Final DataFrame 
        final_cols = ['original row_base', 'original row_test', 'Comparison_Status', 'Change_Logs'] + list(df2.columns)
        final_cols = [col for col in final_cols if col in merged.columns]
        final_df = merged[final_cols]

        # 7. Apply Highlighting logic
        print("Applying visual highlights...")
        def highlight_cells(row):
            styles = pd.Series([''] * len(row), index=row.index)
            status = row.get('Comparison_Status', '')
            logs = str(row.get('Change_Logs', ''))
            
            # 1. Process Length Validation Colors
            length_val = str(row.get(CHECK_COL, "")).strip()
            if CHECK_COL in styles.index and status != 'DELETED ROW':
                if length_val == '12':
                    styles[CHECK_COL] = 'background-color: #99FF99' # Green
                else:
                    styles[CHECK_COL] = 'background-color: #FF9999' # Red
                    
            # 2. Process End Column Validation Colors
            if TEST_END_COL in styles.index and status != 'DELETED ROW':
                if "END MISMATCH" in logs:
                    styles[TEST_END_COL] = 'background-color: #FF9999' # Red
                else:
                    styles[TEST_END_COL] = 'background-color: #99FF99' # Green

            # 3. Process General Modified Columns
            if "Changed:" in logs:
                # Extract the "Changed: col1, col2" part
                changed_part = [part for part in logs.split(" | ") if part.startswith("Changed:")][0]
                changed_cols = changed_part.replace('Changed: ', '').split(', ')
                for col in changed_cols:
                    if col in styles.index:
                        styles[col] = 'background-color: #FF9999' # Red

            # 4. Highlight specific row statuses
            if status == 'DELETED ROW':
                styles[:] = 'background-color: #E0E0E0' # Grey row
            elif status == 'VALIDATION FAILED':
                styles['Comparison_Status'] = 'background-color: #FFFF99' # Yellow warning
            elif status == 'SUCCESS':
                styles['Comparison_Status'] = 'background-color: #99FF99' # Green success
                
            return styles

        styled_df = final_df.style.apply(highlight_cells, axis=1)

        # 8. Save to Excel
        print(f"Saving formatted report to {OUTPUT_FILE}...")
        styled_df.to_excel(OUTPUT_FILE, index=False, engine='openpyxl')
        
        # Print Summary
        print("\n--- FINAL SUMMARY ---")
        print(final_df['Comparison_Status'].value_counts().to_string())
        print("---------------------")
        print("Process Complete!")

    except Exception as e:
        print(f"X Error: {str(e)}")

if __name__ == "__main__":
    run_comparison()
