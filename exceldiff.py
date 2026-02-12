import pandas as pd

# CONFIGURATION
FILE_1 = 'base_file.xlsx'    # Your comparison base
FILE_2 = 'file_with_extras.xlsx' 
OUTPUT = 'difference_report.xlsx'

def find_real_differences():
    print("Loading files...")
    df1 = pd.read_excel(FILE_1)
    df2 = pd.read_excel(FILE_2)

    # We use every column to find a match. 
    # This ignores row position and looks for identical data content.
    print("Finding 'in-between' differences...")
    
    # indicator=True creates a column called '_merge'
    # It will label rows as 'both', 'left_only', or 'right_only'
    diff = pd.merge(df1, df2, how='outer', indicator=True)

    # Rename the labels for clarity
    diff['_merge'] = diff['_merge'].map({
        'left_only': 'Missing in File 2 (Deleted)',
        'right_only': 'Extra in File 2 (The In-Between Rows)',
        'both': 'Matches Perfectly'
    })

    print(f"Saving results to {OUTPUT}...")
    diff.sort_values('_merge', ascending=False).to_excel(OUTPUT, index=False)
    print("Done! Open the output file and filter the '_merge' column.")

if __name__ == "__main__":
    find_real_differences()
