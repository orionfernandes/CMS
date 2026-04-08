import pandas as pd
import os
from datetime import datetime

# ==========================================
# 1. DEFINE YOUR FILE PATHS HERE
# ==========================================
smartsheet = r'p:/DB Summary Report Training/Notation Updates/April 2026/Final Cert Pipeline V3 (6).xlsx'
ca112c     = r'p:/DB Summary Report Training/Notation Updates/April 2026/CA112C Exceptions 20260401_011815.csv'
# ==========================================

def get_col_if_exists(df, col_name):
    """Helper to check if a column exists in the DataFrame."""
    return col_name if col_name in df.columns else None

def build_combined_notation_if_needed(df: pd.DataFrame) -> pd.DataFrame:
    """Processes notation and assignment fields into a single string."""
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

def main():
    # Determine local path for output
    script_directory = os.path.dirname(os.path.abspath(__file__))
    date_str = datetime.now().strftime("%d%m%Y")
    output_filename = f"merged_notation_{date_str}.xlsx"
    output_path = os.path.join(script_directory, output_filename)

    # Check if files exist before processing
    if not os.path.exists(smartsheet) or not os.path.exists(ca112c):
        print(f"Error: One or both files not found./nSmartsheet: {smartsheet}/nCA112C: {ca112c}")
        return

    # Load the files
    print("Loading data...")
    df_smart = pd.read_excel(smartsheet)
    df_ca = pd.read_csv(ca112c)

    # Filter Smartsheet for 'DB' under 'Clients Deal Name'
    df_smart_filtered = df_smart[df_smart['Clients Deal Name'] == 'DB'].copy()

    # Merge/Bounce data
    # Joins 'Unique ID' (Smartsheet) with 'distinct_system_exception_code' (CA112C)
    merged_df = pd.merge(
        df_smart_filtered, 
        df_ca[['distinct_system_exception_code', 'notation', 'document_party_name', 'to_assignment']], 
        left_on='Unique ID', 
        right_on='distinct_system_exception_code', 
        how='inner'
    )

    # Build notation and filter columns
    processed_df = build_combined_notation_if_needed(merged_df)
    final_output = processed_df[['Unique ID', 'combined_notation']].copy()
    
    # Exclude rows where combined_notation is blank or NaN
    final_output = final_output[final_output['combined_notation'].str.strip() != ""]
    final_output = final_output.dropna(subset=['combined_notation'])

    # Export to local folder
    final_output.to_excel(output_path, index=False)
    print(f"Process complete! Output saved to script folder: {output_filename}")

if __name__ == "__main__":
    main()
