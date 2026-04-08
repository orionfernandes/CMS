import pandas as pd
import os
from datetime import datetime

# ==========================================
# USER CONFIGURATION - INPUT FILE PATHS
# ==========================================
# Please update these input paths to where your actual files are located
MAIN_FILE_PATH = "c:/Users/OFernandes/OneDrive - Carrington Mortgage Holdings, LLC/Documents/Notation Test/Final Cert Pipeline V3 (9).xlsx"
PCE_FILE_PATH = "c:/Users/OFernandes/OneDrive - Carrington Mortgage Holdings, LLC/Documents/Notation Test/PCE Pipeline (10).xlsx"
CA112C_FILE_PATH = "c:/Users/OFernandes/OneDrive - Carrington Mortgage Holdings, LLC/Documents/Notation Test/CA112C Exceptions 20260407_011114.csv"
# ==========================================

def get_col_if_exists(df: pd.DataFrame, col_name: str):
    """Helper function: Returns the column name if it exists in the dataframe, else None."""
    return col_name if col_name in df.columns else None

def build_combined_notation_if_needed(df: pd.DataFrame) -> pd.DataFrame:
    """Provided function to combine notation fields."""
    if get_col_if_exists(df, "combined_notation") is not None:
        return df

    notation = get_col_if_exists(df, "notation")
    from_a   = get_col_if_exists(df, "document_party_name")
    to_a     = get_col_if_exists(df, "to_assignment")
    
    if notation is None or from_a is None or to_a is None:
        return df

    df = df.copy()

    def comb(row):
        parts = []
        for col in (notation, from_a, to_a):
            v = row.get(col)
            v = "" if pd.isna(v) else str(v).strip()
            if v:
                parts.append(v)
        return "_".join(parts)

    df["combined_notation"] = df.apply(comb, axis=1)
    return df

def process_file(input_path: str, csv_df: pd.DataFrame, output_filename: str):
    """Processes a single Excel file (main or PCE) against the CSV data."""
    if not os.path.exists(input_path):
        print(f"Error: Could not find the file at {input_path}")
        return

    # 1. Read the Excel file
    df = pd.read_excel(input_path)
    
    # 2. Filter rows where 'Clients Deal Name' is 'DB'
    if 'Clients Deal Name' in df.columns:
        df_filtered = df[df['Clients Deal Name'] == 'DB'].copy()
    else:
        print(f"Warning: 'Clients Deal Name' not found in {input_path}")
        return
        
    # 3 & 4. Isolate required fields and set up 'old_notation'
    required_cols = ['Unique ID', 'Assigned Analyst', 'Doc Notation']
    for col in required_cols:
        if col not in df_filtered.columns:
            print(f"Error: Required column '{col}' missing from {input_path}")
            return
            
    df_filtered = df_filtered.rename(columns={'Doc Notation': 'old_notation'})
    
    # 5. Bounce 'Unique ID' against 'distinct_system_exception_code' to merge CSV fields
    merged_df = pd.merge(
        df_filtered, 
        csv_df[['distinct_system_exception_code', 'notation', 'document_party_name', 'to_assignment']],
        left_on='Unique ID', 
        right_on='distinct_system_exception_code', 
        how='left'
    )
    
    # 6. Apply the function to create 'combined_notation'
    merged_df = build_combined_notation_if_needed(merged_df)
    
    # Rename the output of your function to 'new_notation'
    if 'combined_notation' in merged_df.columns:
        merged_df = merged_df.rename(columns={'combined_notation': 'new_notation'})
    else:
        merged_df['new_notation'] = "" # Fallback if merging yielded no data
        
    # 7. Compare old and new notation to find changes, EXCLUDING blank new_notations
    # We cast to string and fill empty values to ensure a clean comparison
    merged_df['old_notation'] = merged_df['old_notation'].fillna('').astype(str).str.strip()
    merged_df['new_notation'] = merged_df['new_notation'].fillna('').astype(str).str.strip()
    
    # Filter for rows where notations are different AND new_notation is not blank
    changed_df = merged_df[
        (merged_df['old_notation'] != merged_df['new_notation']) & 
        (merged_df['new_notation'] != "")
    ]
    
    # Select only the columns you requested for the final sheet
    final_columns = ['Unique ID', 'Assigned Analyst', 'old_notation', 'new_notation']
    final_df = changed_df[final_columns]
    
    # Output directly to the local folder
    final_df.to_excel(output_filename, index=False)
    print(f"Successfully processed {input_path}. Found {len(final_df)} actionable notation changes. Saved locally as '{output_filename}'.")

def main():
    print("Starting data processing...")
    
    # Get today's date formatted as MMDDYYYY for the output files
    current_date = datetime.now().strftime("%m%d%Y")
    
    # Read the CSV file once since both main and PCE use it
    if not os.path.exists(CA112C_FILE_PATH):
        print(f"Error: Could not find the CSV file at {CA112C_FILE_PATH}")
        return
        
    print(f"Reading CSV file: {CA112C_FILE_PATH}...")
    ca112c_df = pd.read_csv(CA112C_FILE_PATH)

    # Process Main file and output locally
    print("/n--- Processing 'main' file ---")
    local_output_main = f"main_notations_{current_date}.xlsx"
    process_file(MAIN_FILE_PATH, ca112c_df, local_output_main)
    
    # Process PCE file and output locally
    print("/n--- Processing 'PCE' file ---")
    local_output_pce = f"PCE_notations_{current_date}.xlsx"
    process_file(PCE_FILE_PATH, ca112c_df, local_output_pce)
    
    print("/nAll processing is complete!")

if __name__ == "__main__":
    main()
