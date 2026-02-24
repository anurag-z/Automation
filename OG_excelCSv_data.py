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

# --- Define the range to compare (0-based indexing) ---
START_RECORD = 0
END_RECORD = 500  # Set to None to process to the very end: END_RECORD = None

def run_comparison():
    try:
        print("Loading CSV data...")
        df1 = pd.read_csv(FILE_1_BASE, dtype=str)
        df2 = pd.read_csv(FILE_2_NEW, dtype=str)

        # --- NEW: Capture original row numbers before any filtering ---
        # Add 2 because pandas is 0-indexed and CSVs have 1 header row
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
        match_cols = ["Area", "Screen Name", "Screen Number", "Field Name", "Level", "Row", "Column"]
        for col in match_cols:
            if col in df1.columns and col in df2.columns:
                df1[col] = df1[col].fillna("")
                df2[col] = df2[col].fillna("")

        # --- NEW: Capture original row numbers for test data ---
        df2_with_idx = df2.copy()
        df2_with_idx['original row_test'] = df2_with_idx.index + 2 
        
        # 4. Perform Left Join
        merged = pd.merge(df1, df2_with_idx, on=match_cols, how='left', suffixes=('_base', ''))

        # 5. Row processing logic
        def process_row(row):
            status = "COMMON"
            change_logs = ""
            
            # Check for DELETED (If there is no test row number, it didn't match)
            if pd.isna(row.get('original row_test')):
                return pd.Series(["DELETED ROW", ""])

            # Check for MODIFIED
            changes = []
            for col in df2.columns:
                base_val = str(row.get(f"{col}_base", "")).strip()
                new_val = str(row.get(col, "")).strip()
                
                if pd.notna(row.get(f"{col}_base")) and base_val != new_val:
                    changes.append(col)
            
            if changes:
                status = "MODIFIED"
                change_logs = f"Changed: {', '.join(changes)}"
            
            # Validation Check (overwrites status if invalid length, but preserves change_logs)
            if str(row.get(CHECK_COL, "")).strip() != "10":
                status = "INVALID LENGTH"
                
            return pd.Series([status, change_logs])

        print("Checking for differences...")
        merged[['Comparison_Status', 'Change_Logs']] = merged.apply(process_row, axis=1)

        # 6. Build Final DataFrame (Ensure our new columns are first!)
        final_cols = ['original row_base', 'original row_test', 'Comparison_Status', 'Change_Logs'] + list(df2.columns)
        final_cols = [col for col in final_cols if col in merged.columns]
        final_df = merged[final_cols]

        # 7. Apply Highlighting logic
        print("Applying visual highlights...")
        def highlight_cells(row):
            styles = pd.Series([''] * len(row), index=row.index)
            status = row.get('Comparison_Status', '')
            logs = str(row.get('Change_Logs', ''))
            
            # Get the current value of the Length column from the 2nd Excel
            length_val = str(row.get(CHECK_COL, "")).strip()
            
            # 1. Process all columns that registered a change
            if logs.startswith('Changed: '):
                changed_cols = logs.replace('Changed: ', '').split(', ')
                for col in changed_cols:
                    if col in styles.index:
                        if col == CHECK_COL:
                            # If 'Length' changed, color it Green if it is 12, otherwise Red
                            if length_val == '12':
                                styles[col] = 'background-color: #99FF99' # Light Green
                            else:
                                styles[col] = 'background-color: #FF9999' # Light Red
                        else:
                            # Any other changed column gets colored Red
                            styles[col] = 'background-color: #FF9999' # Light Red
                            
            # 2. What if 'Length' failed to update at all? 
            # If it didn't change, it won't be in the logs, but it's still wrong.
            if CHECK_COL in styles.index and length_val != '12' and status != 'DELETED ROW':
                styles[CHECK_COL] = 'background-color: #FF9999' # Force Red

            # 3. Highlight specific row statuses
            if status == 'DELETED ROW':
                styles[:] = 'background-color: #E0E0E0' # Grey row
            elif status == 'INVALID LENGTH':
                styles['Comparison_Status'] = 'background-color: #FFFF99' # Yellow status cell
                
            return styles

        styled_df = final_df.style.apply(highlight_cells, axis=1)

        # 8. Save to Excel
        print(f"Saving formatted report to {OUTPUT_FILE}...")
        styled_df.to_excel(OUTPUT_FILE, index=False, engine='openpyxl')
        print("Process Complete!")

    except Exception as e:
        print(f"X Error: {str(e)}")

if __name__ == "__main__":
    run_comparison()
