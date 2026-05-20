import smartsheet
import pandas as pd
import datetime

# ==========================================================
# 1. CONFIGURATION
# ==========================================================
SMARTSHEET_ACCESS_TOKEN = 'API Token'

# Main Sheet IDs
PIPELINE_ID = 'PIPELINE ID'
ADDS_ID = 'ADDS ID"
CLEARS_ID = 'CLEARS ID'

# Files
DATA_FILE = 'test_import.xlsx'
MAPPING_FILE = 'Smartsheet_Column_IDs.xlsx'
CLEARS_FILE = 'Clears.xlsx'

# Initialize Client
smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
smartsheet_client.errors_as_exceptions(True)

def get_mapping(sheet_name):
    """Dynamically loads mapping from a specific Excel tab (PIPELINE, ADDS, or CLEARS)."""
    try:
        df = pd.read_excel(MAPPING_FILE, sheet_name=sheet_name)
        return dict(zip(df['Column Name'], df['Column ID']))
    except Exception as e:
        print(f"!!! Error loading mapping tab '{sheet_name}': {e}")
        return {}

def process_clears():
    print("--- Starting Clears Processing ---")
    
    # Load mapping for the Pipeline sheet to identify target columns
    col_map = get_mapping('PIPELINE')
    
    try:
        # Load the Unique IDs and dates from the Clears file
        clears_df = pd.read_excel(CLEARS_FILE)
        
        # Download the current state of the Pipeline sheet
        pipeline_sheet = smartsheet_client.Sheets.get_sheet(PIPELINE_ID)
        
        # Get IDs from your Excel mapping tab
        uid_col = col_map.get("Unique ID")
        status_col = col_map.get("Document Exception Status")
        date_col = col_map.get("Cleared Date")

        if not all([uid_col, status_col, date_col]):
            print("Error: Mapping file (PIPELINE tab) is missing IDs for Unique ID, Status, or Date.")
            return

        # Map existing Pipeline rows strictly by their Unique ID value
        pipeline_rows = {}
        for row in pipeline_sheet.rows:
            for cell in row.cells:
                if cell.column_id == uid_col:
                    pipeline_rows[str(cell.value)] = row

        rows_to_update = []
        row_ids_to_move = []

        # Process each row in the Clears Excel file based on Unique ID
        for _, excel_row in clears_df.iterrows():
            excel_uid = str(excel_row['Unique ID'])
            
            if excel_uid in pipeline_rows:
                smartsheet_row = pipeline_rows[excel_uid]
                
                # Format the Cleared Date for Smartsheet (YYYY-MM-DD)
                cleared_date = ""
                if pd.notna(excel_row['Cleared Date']):
                    cleared_date = pd.to_datetime(excel_row['Cleared Date']).strftime('%Y-%m-%d')

                # Prepare the update: Set Status to 'Cleared' and set the Date
                new_row = smartsheet.models.Row()
                new_row.id = smartsheet_row.id
                new_row.cells.append({'column_id': int(status_col), 'value': 'Cleared'})
                new_row.cells.append({'column_id': int(date_col), 'value': cleared_date})
                
                rows_to_update.append(new_row)
                row_ids_to_move.append(smartsheet_row.id)

        if rows_to_update:
            # 1. Apply the updates to the Pipeline sheet first
            print(f"Updating {len(rows_to_update)} rows in Pipeline...")
            smartsheet_client.Sheets.update_rows(PIPELINE_ID, rows_to_update)
            
            # 2. Move the rows to the Clears sheet
            print(f"Moving {len(row_ids_to_move)} rows to CLEARS sheet (ID: {CLEARS_ID})...")
            
            # --- THE FIX: Dictionary Initialization based on your provided code ---
            directive = smartsheet_client.models.CopyOrMoveRowDirective({
                "row_ids": row_ids_to_move,
                "to": smartsheet_client.models.CopyOrMoveRowDestination({
                    "sheet_id": CLEARS_ID
                })
            })
            
            # Execute the move
            response = smartsheet_client.Sheets.move_rows(PIPELINE_ID, directive)
            print(f"Success! Server Message: {getattr(response, 'message', 'OK')}")
        else:
            print("No matching Unique IDs found in the Pipeline sheet.")
            
    except Exception as e:
        print(f"!!! Error in process_clears: {e}")

if __name__ == "__main__":
    process_clears()
