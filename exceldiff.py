import pandas as pd
import gc

# CONFIGURATION
FILE_1 = 'base_file.csv'         # Your comparison base
FILE_2 = 'file_with_extras.csv' 
OUTPUT = 'difference_report.csv'

def find_csv_differences():
    try:
        print("🚀 Loading CSV files...")
        # low_memory=False helps avoid mixed data type warnings in large files
        df1 = pd.read_csv(FILE_1, low_memory=False)
        df2 = pd.read_csv(FILE_2, low_memory=False)

        print("🧹 Cleaning duplicates to ensure alignment...")
        df1 = df1.drop_duplicates()
        df2 = df2.drop_duplicates()

        print("🔄 Comparing all columns (Fuzzy Alignment)...")
        # By not specifying an 'on' column, pandas compares all shared columns
        diff = pd.merge(df1, df2, how='outer', indicator='Result_Status')

        # Clean up memory
        del df1
        del df2
        gc.collect()

        # Mapping the labels for easier reading
        status_map = {
            'left_only': 'Row Missing in File 2',
            'right_only': 'Extra In-Between Row in File 2',
            'both': 'Exact Match'
        }
        diff['Result_Status'] = diff['Result_Status'].map(status_map)

        print(f"💾 Saving results to {OUTPUT}...")
        diff.to_csv(OUTPUT, index=False)
        
        print("\n✅ DONE!")
        print(f"Matches: {len(diff[diff['Result_Status'] == 'Exact Match'])}")
        print(f"Extras: {len(diff[diff['Result_Status'] == 'Extra In-Between Row in File 2'])}")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    find_csv_differences()
