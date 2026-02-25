import pandas as pd

# --- Configuration ---
FILE_1_BASE = r"C:\Test\FieldAnalysis_Combined_1040_18_Feb.csv"
FILE_2_NEW = r"C:\Test\FieldAnalysis_Combined_1040_test.csv"
OUTPUT_FILE = r"C:\Test\Python_excel\difference_report_Full_Compare.xlsx"

CHECK_COL = "Length"

def run_comparison():
    try:
        print("Loading CSV data (No filters applied)...")
        # Read files and fill nulls with empty strings to make comparison clean
        df1 = pd.read_csv(FILE_1_BASE, dtype=str).fillna("")
        df2 = pd.read_csv(FILE_2_NEW, dtype=str).fillna("")

        # 1. Capture original row numbers (adding 2 for Excel header & 0-indexing)
        df1['original row_base'] = df1.index + 2
        df2['original row_test'] = df2.index + 2

        # 2. Define match columns
        match_cols = ["Area", "Screen Name", "Screen Number", "Field Name", "Level", "Row", "Column"]
        
        print(f"Comparing {len(df1)} Base records against {len(df2)} Test records...")

        # 3. Perform Left Join (This replaces the C# nested 'for' loop instantly)
        merged = pd.merge(df1, df2, on=match_cols, how='left', suffixes=('_base', ''))

        # 4. Row processing logic
        def process_row(row):
            status = "COMMON"
            change_logs = ""
            
            # If original row_test is empty, it means no match was found in df2
            if pd.isna(row.get('original row_test')) or row.get('original row_test') == "":
                return pd.Series(["DELETED ROW", ""])

            # Check for MODIFIED columns
            changes = []
            for col in df2.columns:
                # We don't compare the match columns or the tracking column
                if col in match_cols or col == 'original row_test':
                    continue
                
                # Get base and new values
                base_val = str(row.get(f"{col}_base", "")).strip()
                new_val = str(row.get(col, "")).strip()
                
                # If they are different, log the change
                if base_val != new_val:
                    changes.append(col)
            
            if changes:
                status = "MODIFIED"
                change_logs = f"Changed: {', '.join(changes)}"
            
            # Validation Check: If Length is not exactly 12, flag as INVALID LENGTH
            if str(row.get(CHECK_COL, "")).strip() != "12":
                status = "INVALID LENGTH"
                
            return pd.Series([status, change_logs])

        print("Checking for differences...")
        merged[['Comparison_Status', 'Change_Logs']] = merged.apply(process_row, axis=1)

        # 5. Build Final DataFrame 
        # Ensure our tracking and status columns are perfectly ordered at the front
        final_cols = ['original row_base', 'original row_test', 'Comparison_Status', 'Change_Logs'] + list(df2.columns)
        
        # Remove the tracking column from the end since we moved it to the front
        final_cols = [col for col in final_cols if col in merged.columns]
        # De-duplicate column list just in case
        final_cols = list(dict.fromkeys(final_cols)) 
        
        final_df = merged[final_cols]

        # 6. Apply Highlighting logic
        print("Applying visual highlights...")
        def highlight_cells(row):
            styles = pd.Series([''] * len(row), index=row.index)
            status = row.get('Comparison_Status', '')
            logs = str(row.get('Change_Logs', ''))
            
            # Get the current value of the Length column
            length_val = str(row.get(CHECK_COL, "")).strip()
            
            # Process changed columns
            if logs.startswith('Changed: '):
                changed_cols = logs.replace('Changed: ', '').split(', ')
                for col in changed_cols:
                    if col in styles.index:
                        if col == CHECK_COL:
                            # Length changed to 12 = Green. Changed to anything else = Red.
                            if length_val == '12':
                                styles[col] = 'background-color: #99FF99' # Light Green
                            else:
                                styles[col] = 'background-color: #FF9999' # Light Red
                        else:
                            # Other modified columns = Red
                            styles[col] = 'background-color: #FF9999' # Light Red
                            
            # Check Length even if it didn't change 
            if CHECK_COL in styles.index and length_val != '12' and status != 'DELETED ROW':
                styles[CHECK_COL] = 'background-color: #FF9999' # Force Red

            # Row/Status Specific Highlights
            if status == 'DELETED ROW':
                styles[:] = 'background-color: #E0E0E0' # Grey row
            elif status == 'INVALID LENGTH':
                styles['Comparison_Status'] = 'background-color: #FFFF99' # Yellow status cell
                
            return styles

        styled_df = final_df.style.apply(highlight_cells, axis=1)

        # 7. Save to Excel
        print(f"Saving formatted report to {OUTPUT_FILE}...")
        styled_df.to_excel(OUTPUT_FILE, index=False, engine='openpyxl')
        print("Process Complete!")

    except Exception as e:
        print(f"X Error: {str(e)}")

if __name__ == "__main__":
    run_comparison()
