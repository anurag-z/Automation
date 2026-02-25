import pandas as pd

# --- Configuration ---
PREVIOUS_OUTPUT_FILE = r"C:\Test\Python_excel\difference_report_Full_Compare.xlsx"
FILE_2_NEW = r"C:\Test\FieldAnalysis_Combined_1040_test.csv"
OUTPUT_FILE = r"C:\Test\Python_excel\Deleted_Row_Analysis.xlsx"

def investigate_deleted_rows():
    try:
        print("Loading 'DELETED ROWS' from previous output...")
        
        # FIX APPLIED HERE: Read as object, fill nulls, then convert all to string
        df_diff = pd.read_excel(PREVIOUS_OUTPUT_FILE, dtype=object).fillna("").astype(str)
        
        df_deleted = df_diff[df_diff['Comparison_Status'] == 'DELETED ROW'].copy()
        
        if len(df_deleted) == 0:
            print("Great news! There are 0 DELETED ROWS in your report. Nothing to investigate.")
            return

        print(f"Found {len(df_deleted)} deleted rows to investigate.")

        print("Loading Original Test File to search for matches...")
        
        # FIX APPLIED HERE ALSO: Read as object, fill nulls, then convert all to string
        df_test = pd.read_csv(FILE_2_NEW, dtype=object).fillna("").astype(str)

        relaxed_match_cols = ["Area", "Screen Name", "Screen Number", "Field Name", "Field Number"]
        
        print(f"Attempting to find rows using only: {relaxed_match_cols}...")

        analysis_results = []

        for index, deleted_row in df_deleted.iterrows():
            
            condition = pd.Series([True] * len(df_test))
            for col in relaxed_match_cols:
                val = str(deleted_row.get(col, "")).strip()
                if not val and f"{col}_base" in deleted_row:
                    val = str(deleted_row.get(f"{col}_base", "")).strip()
                
                if col in df_test.columns:
                    condition &= (df_test[col].str.strip() == val)
            
            matches = df_test[condition]
            result_row = deleted_row.copy()
            
            base_length = str(deleted_row.get('Length_base', deleted_row.get('Length', ''))).strip()
            result_row['Base_Length'] = base_length

            if len(matches) > 0:
                match = matches.iloc[0] 
                
                result_row['Investigation_Status'] = "FOUND (MOVED/CHANGED)"
                result_row['original row_test'] = str(match.name + 2)  
                result_row['New_Length'] = str(match.get('Length', '')).strip() 
                
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
                result_row['Investigation_Status'] = "TRULY DELETED"
                result_row['Reason_for_Mismatch'] = "Field not found in Test File (even with relaxed match)"
                result_row['original row_test'] = "N/A"
                result_row['New_Length'] = "N/A"
            
            analysis_results.append(result_row)

        df_analysis = pd.DataFrame(analysis_results)
        
        cols_to_show = [
            'original row_base', 
            'original row_test', 
            'Investigation_Status', 
            'Reason_for_Mismatch',
            'Base_Length',
            'New_Length'
        ] + relaxed_match_cols + ['Level', 'Row', 'Column']
        
        extra_cols = [c for c in df_analysis.columns if c.startswith("New_") and c not in cols_to_show]
        cols_to_show += extra_cols
        
        final_cols = [c for c in cols_to_show if c in df_analysis.columns]
        df_final = df_analysis[final_cols]
        
        print(f"Saving investigation report to {OUTPUT_FILE}...")
        
        def highlight_status(row):
            styles = pd.Series([''] * len(row), index=row.index)
            if row['Investigation_Status'] == "TRULY DELETED":
                styles[:] = 'background-color: #FF9999'
            else:
                styles[:] = 'background-color: #99FF99'
            return styles
                
        df_final.style.apply(highlight_status, axis=1).to_excel(OUTPUT_FILE, index=False, engine='openpyxl')
        print("Done! Process Complete.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    investigate_deleted_rows()
