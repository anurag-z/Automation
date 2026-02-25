import pandas as pd

# --- Configuration ---
FILE_1_BASE = r"C:\Test\FieldAnalysis_Combined_1040_18_Feb.csv"
FILE_2_NEW = r"C:\Test\FieldAnalysis_Combined_1040_test.csv"
OUTPUT_FILE = r"C:\Test\Python_excel\difference_report_Full_Compare.xlsx"

# --- Define the range to compare (0-based indexing) ---
# Example: 0 to 500. To process everything, set END_RECORD = None
START_RECORD = 0
END_RECORD = 500 

def run_comparison():
    try:
        print("Loading CSV data (No filters applied)...")
        # Read files and fill nulls with empty strings to make comparison clean
        df1 = pd.read_csv(FILE_1_BASE, dtype=str).fillna("")
        df2 = pd.read_csv(FILE_2_NEW, dtype=str).fillna("")

        # 1. Capture original row numbers before slicing
        df1['original row_base'] = df1.index + 2
        df2['original row_test'] = df2.index + 2

        # 2. Apply the Range Slice to the Base file
        total_base = len(df1)
        actual_end = min(END_RECORD if END_RECORD is not None else total_base, total_base)
        
        # Slice df1 based on the start and end records
        df1 = df1.iloc[START_RECORD:actual_end]

        # 3. Define match columns
        match_cols = ["Area", "Screen Name", "Screen Number", "Field Name", "Level", "Row", "Column"]
        
        print(f"Comparing {len(df1)} Base records (Index {START_RECORD} to {actual_end - 1}) against {len(df2)} Test records...")

        # 4. Perform Left Join
        merged = pd.merge(df1, df2, on=match_cols, how='left', suffixes=('_base', ''))

        # 5. Row processing logic - Purely checking for diffs
        def process_row(row):
            status = "COMMON"
            change_logs = ""
            
            # If original row_test is empty, no match was found in df2
            if pd.isna(row.get('original row_test')) or row.get('original row_test') == "":
                return pd.Series(["DELETED ROW", ""])

            # Check for MODIFIED columns
            changes = []
            for col in df2.columns:
                # Skip tracking/match columns
                if col in match_cols or col == 'original row_test':
                    continue
                
                # Compare values
                base_val = str(row.get(f"{col}_base", "")).strip()
                new_val = str(row.get(col, "")).strip()
                
                if base_val != new_val:
                    changes.append(col)
            
            if changes:
                status = "MODIFIED"
                change_logs = f"Changed: {', '.join(changes)}"
                
            return pd.Series([status, change_logs])

        print("Checking for differences...")
        merged[['Comparison_Status', 'Change_Logs']] = merged.apply(process_row, axis=1)

        # 6. Build Final DataFrame 
        final_cols = ['original row_base', 'original row_test', 'Comparison_Status', 'Change_Logs'] + list(df2.columns)
        
        # Clean up column list
        final_cols = [col for col in final_cols if col in merged.columns]
        final_cols = list(dict.fromkeys(final_cols)) 
        
        final_df = merged[final_cols]

        # 7. Apply Highlighting logic
        print("Applying visual highlights...")
        def highlight_cells(row):
            styles = pd.Series([''] * len(row), index=row.index)
            status = row.get('Comparison_Status', '')
            logs = str(row.get('Change_Logs', ''))
            
            # Highlight any changed columns in Red
            if logs.startswith('Changed: '):
                changed_cols = logs.replace('Changed: ', '').split(', ')
                for col in changed_cols:
                    if col in styles.index:
                        styles[col] = 'background-color: #FF9999' # Light Red
                        
            # Row/Status Specific Highlights
            if status == 'DELETED ROW':
                styles[:] = 'background-color: #E0E0E0' # Grey row
            elif status == 'MODIFIED':
                styles['Comparison_Status'] = 'background-color: #FFFF99' # Yellow status cell for visibility
                
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
