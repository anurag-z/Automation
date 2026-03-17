import pandas as pd

# 1. Load the Excel file
file_path = 'your_file_name.xlsx'
excel_file = pd.ExcelFile(file_path)

# 2. List of your sheet names
sheet_names = ['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4']

# 3. Read and combine them
all_data = []
for sheet in sheet_names:
    print(f"Reading {sheet}...")
    df = pd.read_excel(excel_file, sheet_name=sheet)
    all_data.append(df)

# 4. Merge (Stack) everything together
print("Merging data...")
merged_df = pd.concat(all_data, ignore_index=True)

# 5. Save to CSV (Excel sheets won't support 40 lakh rows)
output_file = 'merged_data_40_lakh.csv'
merged_df.to_csv(output_file, index=False)

print(f"Done! Your merged file is saved as {output_file}")
