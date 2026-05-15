import smartsheet
import pandas as pd
import datetime

# ==========================================================
# 1. CONFIGURATION - Update these with your real info
# ==========================================================
SMARTSHEET_ACCESS_TOKEN = 'API KEY'
SHEET_ID = 'Sheet ID'
DATA_FILE = 'test_import.xlsx'
MAPPING_FILE = 'Smartsheet_Column_IDs.xlsx'

# 2. INITIALIZE THE CLIENT
# This MUST be defined before the function is called
smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

def sync_excel_to_smartsheet():
    print("--- Starting Sync Process ---")
    try:
        # 3. Load the Mapping (Excel to Column IDs)
        print(f"Reading mapping file: {MAPPING_FILE}...")
        mapping_df = pd.read_excel(MAPPING_FILE)
        col_map = dict(zip(mapping_df['Column Name'], mapping_df['Column ID']))
        
        # 4. Load your actual data
        print(f"Reading data file: {DATA_FILE}...")
        data_df = pd.read_excel(DATA_FILE)
        
        rows_to_add = []
        
        for index, row in data_df.iterrows():
            new_row = smartsheet.models.Row()
            new_row.to_bottom = True
            
            for col_name, value in row.items():
                if col_name in col_map:
                    cell_value = value
                    
                    # Logic to handle Dates (Smartsheet needs YYYY-MM-DD)
                    is_date_type = isinstance(value, (pd.Timestamp, datetime.datetime))
                    is_date_column = 'date' in str(col_name).lower()

                    if is_date_type or is_date_column:
                        if pd.notna(value):
                            cell_value = pd.to_datetime(value).strftime('%Y-%m-%d')
                        else:
                            cell_value = ""
                    
                    # Handle empty Excel cells
                    elif pd.isna(value):
                        cell_value = ""

                    new_row.cells.append({
                        'column_id': int(col_map[col_name]),
                        'value': cell_value
                    })
            
            rows_to_add.append(new_row)

        # 5. Send the batch
        if rows_to_add:
            print(f"Sending {len(rows_to_add)} rows to Smartsheet...")
            # We use the smartsheet_client defined above
            response = smartsheet_client.Sheets.add_rows(SHEET_ID, rows_to_add)
            print(f"Success! Server Message: {response.message}")
        else:
            print("Warning: No rows were processed. Check your Excel headers.")

    except smartsheet.exceptions.ApiError as e:
        print(f"!!! Smartsheet API Error: {e.error.result.message}")
    except Exception as e:
        print(f"!!! Python Error: {e}")

# 6. RUN THE SCRIPT
if __name__ == "__main__":
    sync_excel_to_smartsheet()
