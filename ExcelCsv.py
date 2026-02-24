import pandas as pd

# --- Changed to .csv extensions ---
FILE_1_BASE = r"C:\Test\FieldAnalysis_Combined_1040_18_Feb.csv"
FILE_2_NEW = r"C:\Test\FieldAnalysis_Combined_1040_test.csv"
OUTPUT_FILE = r"C:\Test\difference_report_Full_Compare.csv"

FILTERS = {
    "Safe to Extend": ["Yes"],
    "Action Required": ["Move and Extend (Non-Group)", "Extend Only (Non-Group)"]
}

CHECK_COL = "Length"

def run_comparison():
    try:
        print("Loading and filtering CSV data...")
        
        # 1. Read CSV files
        # Using dtype=str ensures pandas doesn't accidentally drop leading zeros (like "001" to "1")
        df1 = pd.read_csv(FILE_1_BASE, dtype=str)
        df2 = pd.read_csv(FILE_2_NEW, dtype=str)

        # 2. Apply Filters to df1 (Base)
        for key, values in FILTERS.items():
            if key in df1.columns:
                df1 = df1[df1[key].isin([str(v) for v in values])]

        # 3. Define match columns
        match_cols = ["Area", "Screen Name", "Screen Number", "Field Name", "Level", "Row", "Column"]
        
        # Handle missing match columns cleanly
        for col in match_cols:
            if col in df1.columns and col in df2.columns:
                df1[col] = df1[col].fillna("")
                df2[col] = df2[col].fillna("")
            else:
                print(f"Warning: Match column '{col}' is missing.")

        # 4. Perform a Left Join to find matches
        df2_with_idx = df2.copy()
        df2_with_idx['Original row'] = df2_with_idx.index + 2 
        
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
            
            # Check for DELETED
            if pd.isna(row.get('Original row')):
                return pd.Series(["DELETED ROW", "", ""])

            # Check for MODIFIED
            changes = []
            for col in df2.columns:
                base_val = str(row.get(f"{col}_base", "")).strip()
                new_val = str(row.get(col, "")).strip()
                
                # Check if base value existed and is different
                if pd.notna(row.get(f"{col}_base")) and base_val != new_val:
                    changes.append(col)
            
            if changes:
                status = "MODIFIED"
                change_logs = f"Changed: {', '.join(changes)}"
            
            # Validation Check
            if str(row.get(CHECK_COL, "")).strip() != "10":
                status = "INVALID LENGTH"
                
            return pd.Series([status, change_logs, row['Original row']])

        # 5. Apply the logic
        print("Comparing rows...")
        merged[['Comparison_Status', 'Change_Logs', 'Original row']] = merged.apply(process_row, axis=1)

        # 6. Reorder columns
        final_cols = ['Original row', 'Comparison_Status', 'Change_Logs'] + list(df2.columns)
        final_cols = [col for col in final_cols if col in merged.columns]
        final_df = merged[final_cols]

        # 7. Save to CSV
        print(f"Saving report to {OUTPUT_FILE}...")
        # index=False prevents pandas from writing row numbers as the first column
        final_df.to_csv(OUTPUT_FILE, index=False)
        print("Process Complete!")

    except Exception as e:
        print(f"X Error: {str(e)}")

if __name__ == "__main__":
    run_comparison()


import pandas as pd

# --- Configuration ---
FILE_1_BASE = r"C:\Test\FieldAnalysis_Combined_1040_18_Feb.csv"
FILE_2_NEW = r"C:\Test\FieldAnalysis_Combined_1040_test.csv"
OUTPUT_FILE = r"C:\Test\difference_report_Full_Compare.xlsx"

FILTERS = {
    "Safe to Extend": ["Yes"],
    "Action Required": ["Move and Extend (Non-Group)", "Extend Only (Non-Group)"]
}
CHECK_COL = "Length"

# --- NEW: Define the range to compare (0-based indexing) ---
# Example: 0 to 500 will process the first 500 records AFTER filtering.
START_RECORD = 0
END_RECORD = 500  # Set to None if you want to process to the very end: END_RECORD = None

def run_comparison():
    try:
        print("Loading CSV data...")
        df1 = pd.read_csv(FILE_1_BASE, dtype=str)
        df2 = pd.read_csv(FILE_2_NEW, dtype=str)

        # 1. Apply Filters to df1
        print("Applying filters...")
        for key, values in FILTERS.items():
            if key in df1.columns:
                df1 = df1[df1[key].isin([str(v) for v in values])]

        # --- NEW: Apply the Range Slice AFTER filtering ---
        total_filtered = len(df1)
        print(f"Total Base records after filtering: {total_filtered}")
        
        # Slice the dataframe. If END_RECORD is larger than total_filtered, pandas handles it safely.
        df1 = df1.iloc[START_RECORD:END_RECORD]
        actual_end = min(END_RECORD if END_RECORD else total_filtered, total_filtered)
        print(f"Comparing a range of {len(df1)} records (Index {START_RECORD} to {actual_end})...")

        # 2. Define match columns and prep for merge
        match_cols = ["Area", "Screen Name", "Screen Number", "Field Name", "Level", "Row", "Column"]
        for col in match_cols:
            if col in df1.columns and col in df2.columns:
                df1[col] = df1[col].fillna("")
                df2[col] = df2[col].fillna("")

        df2_with_idx = df2.copy()
        df2_with_idx['Original row'] = df2_with_idx.index + 2 
        
        # 3. Perform Left Join
        merged = pd.merge(df1, df2_with_idx, on=match_cols, how='left', suffixes=('_base', ''))

        # 4. Row processing logic
        def process_row(row):
            status = "COMMON"
            change_logs = ""
            
            if pd.isna(row.get('Original row')):
                return pd.Series(["DELETED ROW", "", ""])

            changes = []
            for col in df2.columns:
                base_val = str(row.get(f"{col}_base", "")).strip()
                new_val = str(row.get(col, "")).strip()
                
                if pd.notna(row.get(f"{col}_base")) and base_val != new_val:
                    changes.append(col)
            
            if changes:
                status = "MODIFIED"
                change_logs = f"Changed: {', '.join(changes)}"
            
            if str(row.get(CHECK_COL, "")).strip() != "10":
                status = "INVALID LENGTH"
                
            return pd.Series([status, change_logs, row['Original row']])

        print("Checking for differences...")
        merged[['Comparison_Status', 'Change_Logs', 'Original row']] = merged.apply(process_row, axis=1)

        # 5. Build Final DataFrame
        final_cols = ['Original row', 'Comparison_Status', 'Change_Logs'] + list(df2.columns)
        final_cols = [col for col in final_cols if col in merged.columns]
        final_df = merged[final_cols]

        # 6. Apply Highlighting logic
        print("Applying visual highlights...")
        def highlight_cells(row):
            # Create a Series of empty styles mapped to exact column headers
            styles = pd.Series([''] * len(row), index=row.index)
            status = row.get('Comparison_Status', '')
            
            if status == 'MODIFIED':
                logs = str(row.get('Change_Logs', ''))
                if logs.startswith('Changed: '):
                    # Extract the list of columns that changed
                    changed_cols = logs.replace('Changed: ', '').split(', ')
                    for col in changed_cols:
                        if col in styles.index:
                            # Color ONLY the specific modified cell in this row red
                            styles[col] = 'background-color: #FF9999'
                            
            elif status == 'DELETED ROW':
                # Color the entire row grey
                styles[:] = 'background-color: #E0E0E0'
                
            elif status == 'INVALID LENGTH':
                # Color ONLY the status cell yellow
                styles['Comparison_Status'] = 'background-color: #FFFF99'
                
            return styles

        # Apply the style function
        styled_df = final_df.style.apply(highlight_cells, axis=1)

        # 7. Save to Excel
        print(f"Saving formatted report to {OUTPUT_FILE}...")
        # Remember to keep your OUTPUT_FILE as a .xlsx extension!
        styled_df.to_excel(OUTPUT_FILE, index=False, engine='openpyxl')
        print("Process Complete!")

    

    except Exception as e:
        print(f"X Error: {str(e)}")

if __name__ == "__main__":
    run_comparison()
