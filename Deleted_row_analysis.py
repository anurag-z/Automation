import pandas as pd

# --- Configuration ---
# 1. Input the OUTPUT file from the previous script (to get the list of deleted rows)
PREVIOUS_OUTPUT_FILE = r"C:\Test\Python_excel\difference_report_Full_Compare.xlsx"

# 2. Input the ORIGINAL Test CSV (where we look for the missing rows)
FILE_2_NEW = r"C:\Test\FieldAnalysis_Combined_1040_test.csv"

# 3. Output for this analysis
OUTPUT_FILE = r"C:\Test\Python_excel\Deleted_Row_Analysis.xlsx"

def investigate_deleted_rows():
    try:
        print("Loading 'DELETED ROWS' from previous output...")
        df_diff = pd.read_excel(PREVIOUS_OUTPUT_FILE, dtype=str).fillna("")
        
        # Filter for only the rows that were marked as DELETED
        df_deleted = df_diff[df_diff['Comparison_Status'] == 'DELETED ROW'].copy()
        
        if len(df_deleted) == 0:
            print("Great news! There are 0 DELETED ROWS in your report. Nothing to investigate.")
            return

        print(f"Found {len(df_deleted)} deleted rows to investigate.")

        print("Loading Original Test File to search for matches...")
        df_test = pd.read_csv(FILE_2_NEW, dtype=str).fillna("")

        # --- THE RELAXED MATCH STRATEGY ---
        # We assume 'Row', 'Column', and 'Level' might have changed.
        # So we only match on the "Identity" of the field.
        # ADJUST THIS LIST if 'Screen Number' is also changing!
        relaxed_match_cols = ["Area", "Screen Name", "Screen Number", "Field Name"]
        
        print(f"Attempting to find rows using only: {relaxed_match_cols}...")

        # Prepare a list to store results
        analysis_results = []

        # Iterate through each "Deleted" row
        for index, deleted_row in df_deleted.iterrows():
            
            # 1. Build a filter for the Test file based on the Relaxed Columns
            # (This acts like a VLOOKUP on just the 4 identity columns)
            condition = pd.Series([True] * len(df_test))
            for col in relaxed_match_cols:
                val = str(deleted_row[col]).strip()
                condition &= (df_test[col].str.strip() == val)
            
            # 2. Find matches
            matches = df_test[condition]
            
            result_row = deleted_row.copy()
            
            if len(matches) > 0:
                # MATCH FOUND! The row wasn't deleted, it just moved/changed.
                match = matches.iloc[0] # Take the first match found
                
                result_row['Investigation_Status'] = "FOUND (MOVED/CHANGED)"
                
                # Check exactly WHY it failed the original strict match
                changes = []
                strict_cols_to_check = ["Level", "Row", "Column"] # The ones we ignored
                
                for col in strict_cols_to_check:
                    old_val = str(deleted_row[col]).strip()
                    new_val = str(match[col]).strip()
                    if old_val != new_val:
                        changes.append(f"{col}: '{old_val}' -> '{new_val}'")
                        # Add the new value to the report for visibility
                        result_row[f"New_{col}"] = new_val
                
                result_row['Reason_for_Mismatch'] = "; ".join(changes)
                result_row['New_Test_Row_Index'] = match.name + 2 # +2 for Excel row number

            else:
                # STILL NOT FOUND? It was truly deleted.
                result_row['Investigation_Status'] = "TRULY DELETED"
                result_row['Reason_for_Mismatch'] = "Field ID not found in Test File"
            
            analysis_results.append(result_row)

        # --- Create DataFrame and Save ---
        df_analysis = pd.DataFrame(analysis_results)
        
        # Organize columns nicely
        cols_to_show = ['original row_base', 'Investigation_Status', 'Reason_for_Mismatch', 'New_Test_Row_Index'] + relaxed_match_cols + ['Level', 'Row', 'Column']
        # Add any "New_..." columns that were created dynamically
        extra_cols = [c for c in df_analysis.columns if c.startswith("New_") and c != 'New_Test_Row_Index']
        cols_to_show += extra_cols
        
        # Filter to ensure columns exist
        final_cols = [c for c in cols_to_show if c in df_analysis.columns]
        
        df_final = df_analysis[final_cols]
        
        print(f"Saving investigation report to {OUTPUT_FILE}...")
        
        # Highlight logic (Green = Found, Red = Truly Deleted)
        def highlight_status(row):
            styles = [''] * len(row)
            if row['Investigation_Status'] == "TRULY DELETED":
                return ['background-color: #FF9999'] * len(row) # Red
            else:
                return ['background-color: #99FF99'] * len(row) # Green
                
        df_final.style.apply(highlight_status, axis=1).to_excel(OUTPUT_FILE, index=False)
        print("Done! Check the Excel file to see where your data went.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    investigate_deleted_rows()
