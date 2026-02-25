import pandas as pd

# --- Configuration ---
# 1. Input the OUTPUT file from your full comparison script
PREVIOUS_OUTPUT_FILE = r"C:\Test\Python_excel\difference_report_Full_Compare.xlsx"

# 2. Input the ORIGINAL Test CSV
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
        # Added 'Field Number' as requested to ensure a highly accurate search!
        relaxed_match_cols = ["Area", "Screen Name", "Screen Number", "Field Name", "Field Number"]
        
        print(f"Attempting to find rows using only: {relaxed_match_cols}...")

        analysis_results = []

        for index, deleted_row in df_deleted.iterrows():
            
            # 1. Build a filter for the Test file based on the Relaxed Columns
            condition = pd.Series([True] * len(df_test))
            for col in relaxed_match_cols:
                # Get value from the deleted row (checking for '_base' suffix just in case)
                val = str(deleted_row.get(col, "")).strip()
                if not val and f"{col}_base" in deleted_row:
                    val = str(deleted_row.get(f"{col}_base", "")).strip()
                
                # Apply the condition if the column exists in the Test file
                if col in df_test.columns:
                    condition &= (df_test[col].str.strip() == val)
            
            # 2. Find matches
            matches = df_test[condition]
            result_row = deleted_row.copy()
            
            # --- CAPTURE BASE LENGTH ---
            # In the diff file, the original base columns usually get a '_base' suffix
            base_length = str(deleted_row.get('Length_base', deleted_row.get('Length', ''))).strip()
            result_row['Base_Length'] = base_length

            if len(matches) > 0:
                # MATCH FOUND! 
                match = matches.iloc[0] 
                
                result_row['Investigation_Status'] = "FOUND (MOVED/CHANGED)"
                result_row['original row_test'] = match.name + 2  # Excel row number in the test file
                
                # --- CAPTURE NEW LENGTH ---
                result_row['New_Length'] = str(match.get('Length', '')).strip() 
                
                # Check exactly WHY it failed the strict match (e.g., Row or Level changed)
                changes = []
                strict_cols_to_check = ["Level", "Row", "Column"] 
                
                for col in strict_cols_to_check:
                    old_val = str(deleted_row.get(col, "")).strip()
                    if not old_val and f"{col}_base" in deleted_row:
                        old_val = str(deleted_row.get(f"{col}_base", "")).strip()
                        
                    new_val = str(match.get(col, "")).strip()
                    if old_val != new_val:
                        changes.append(f"{col}: '{old_val}' -> '{new_val}'")
                        result_row[f"New_{col}"] = new_val
                
                result_row['Reason_for_Mismatch'] = "; ".join(changes)

            else:
                # STILL NOT FOUND
                result_row['Investigation_Status'] = "TRULY DELETED"
                result_row['Reason_for_Mismatch'] = "Field not found in Test File (even with relaxed match)"
                result_row['original row_test'] = "N/A"
                result_row['New_Length'] = "N/A"
            
            analysis_results.append(result_row)

        # --- Create DataFrame and Save ---
        df_analysis = pd.DataFrame(analysis_results)
        
        # Organize columns to put the new info front and center
        cols_to_show = [
            'original row_base', 
            'original row_test', 
            'Investigation_Status', 
            'Reason_for_Mismatch',
            'Base_Length',
            'New_Length'
        ] + relaxed_match_cols + ['Level', 'Row', 'Column']
        
        # Add any "New_..." columns that were created dynamically
        extra_cols = [c for c in df_analysis.columns if c.startswith("New_") and c not in cols_to_show]
        cols_to_show += extra_cols
        
        # Filter to ensure columns exist in our final dataframe
        final_cols = [c for c in cols_to_show if c in df_analysis.columns]
        
        df_final = df_analysis[final_cols]
        
        print(f"Saving investigation report to {OUTPUT_FILE}...")
        
        # Highlight logic (Green = Found, Red = Truly Deleted)
        def highlight_status(row):
            styles = pd.Series([''] * len(row), index=row.index)
            if row['Investigation_Status'] == "TRULY DELETED":
                styles[:] = 'background-color: #FF9999' # Red
            else:
                styles[:] = 'background-color: #99FF99' # Green
            return styles
                
        df_final.style.apply(highlight_status, axis=1).to_excel(OUTPUT_FILE, index=False, engine='openpyxl')
        print("Done! Process Complete.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    investigate_deleted_rows()
