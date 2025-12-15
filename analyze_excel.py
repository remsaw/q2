import pandas as pd
import openpyxl
from pathlib import Path

# Load the Excel file
file_path = Path(r"c:\Q2\second quarter.xlsx")
print(f"Analyzing: {file_path.name}\n")
print("="*80)

# Get all sheet names
excel_file = pd.ExcelFile(file_path)
print(f"\nSheet Names: {excel_file.sheet_names}")
print(f"Number of Sheets: {len(excel_file.sheet_names)}\n")

# Analyze each sheet
for sheet_name in excel_file.sheet_names:
    print("\n" + "="*80)
    print(f"SHEET: {sheet_name}")
    print("="*80)
    
    # Read the sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    print(f"\nShape: {df.shape[0]} rows × {df.shape[1]} columns")
    print(f"\nColumns: {list(df.columns)}")
    
    print("\n--- First 10 rows ---")
    print(df.head(10))
    
    print("\n--- Data Types ---")
    print(df.dtypes)
    
    print("\n--- Basic Statistics ---")
    print(df.describe(include='all'))
    
    print("\n--- Missing Values ---")
    missing = df.isnull().sum()
    if missing.any():
        print(missing[missing > 0])
    else:
        print("No missing values")
    
    # Check for numeric columns and provide additional insights
    numeric_cols = df.select_dtypes(include=['number']).columns
    if len(numeric_cols) > 0:
        print("\n--- Numeric Column Summary ---")
        for col in numeric_cols:
            print(f"\n{col}:")
            print(f"  Total: {df[col].sum()}")
            print(f"  Mean: {df[col].mean():.2f}")
            print(f"  Min: {df[col].min()}")
            print(f"  Max: {df[col].max()}")

print("\n" + "="*80)
print("Analysis Complete!")
print("="*80)
