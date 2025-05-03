import pandas as pd

def load_data(file_path, lead_date_col, lead_val_col, target_date_col, target_val_col, header_row, sheet_name):
    """Loads data from an Excel sheet, creates separate Series for leading and target indicators.

    Args:
        file_path (str): Path to the Excel file.
        lead_date_col (str): Column name for the leading series' date.
        lead_val_col (str): Column name for the leading series' value.
        target_date_col (str): Column name for the target series' date.
        target_val_col (str): Column name for the target series' value.
        header_row (int): 0-indexed row number containing column headers.
        sheet_name (str or int): Name or index of the sheet to load.

    Returns:
        tuple: A tuple containing (pd.Series, pd.Series) for the leading and target series,
               indexed by their respective datetime columns and values renamed.
               Returns (None, None) if loading fails or columns are missing.
    """
    try:
        # Load the entire sheet first
        df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
        print(f"Successfully loaded sheet '{sheet_name}' from {file_path}")
        print(f"Raw columns: {df_raw.columns.tolist()}")

        # --- Create Leading Series ---
        lead_series_raw = None
        if lead_date_col in df_raw.columns and lead_val_col in df_raw.columns:
            lead_df = df_raw[[lead_date_col, lead_val_col]].copy()
            # Convert date column to datetime, coercing errors to NaT
            lead_df[lead_date_col] = pd.to_datetime(lead_df[lead_date_col], errors='coerce')
            # Drop rows where date conversion failed or value is NA before setting index
            lead_df = lead_df.dropna(subset=[lead_date_col, lead_val_col])
            if not lead_df.empty:
                lead_df = lead_df.set_index(lead_date_col)
                lead_series_raw = lead_df[lead_val_col].rename('Leading')
                lead_series_raw.index.name = 'Date' # Standardize index name
                print(f"Created leading series '{lead_val_col}' indexed by '{lead_date_col}'. Length: {len(lead_series_raw)}")
            else:
                 print(f"Warning: No valid data found for leading series ('{lead_date_col}', '{lead_val_col}') after date conversion/NA drop.")
        else:
            print(f"Error: Required leading columns ('{lead_date_col}', '{lead_val_col}') not found in sheet '{sheet_name}'.")

        # --- Create Target Series ---
        target_series_raw = None
        if target_date_col in df_raw.columns and target_val_col in df_raw.columns:
            target_df = df_raw[[target_date_col, target_val_col]].copy()
            # Convert date column to datetime, coercing errors to NaT
            target_df[target_date_col] = pd.to_datetime(target_df[target_date_col], errors='coerce')
             # Drop rows where date conversion failed or value is NA before setting index
            target_df = target_df.dropna(subset=[target_date_col, target_val_col])
            if not target_df.empty:
                target_df = target_df.set_index(target_date_col)
                target_series_raw = target_df[target_val_col].rename('Target')
                target_series_raw.index.name = 'Date' # Standardize index name
                print(f"Created target series '{target_val_col}' indexed by '{target_date_col}'. Length: {len(target_series_raw)}")
            else:
                 print(f"Warning: No valid data found for target series ('{target_date_col}', '{target_val_col}') after date conversion/NA drop.")
        else:
            print(f"Error: Required target columns ('{target_date_col}', '{target_val_col}') not found in sheet '{sheet_name}'.")

        # Return None, None if either series creation failed
        if lead_series_raw is None or target_series_raw is None:
             print("Error: Failed to create one or both series.")
             return None, None

        return lead_series_raw, target_series_raw

    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
        return None, None
    except Exception as e:
        print(f"Error loading data from {file_path}, sheet '{sheet_name}': {e}")
        return None, None

# Example Usage (for testing purposes, will be called from main.py)
# if __name__ == '__main__':
#     # Create a dummy excel file for testing
#     dummy_data = {
#         'MonthYear': ['01/22', '02/22', '03/22', '04/22', '05/22', '06/22'],
#         'IndicatorA': [10, 12, 11, 13, None, 14],
#         'IndicatorB': [100, 105, 103, 106, 109, 110],
#         'OtherData': ['x', 'y', 'z', 'a', 'b', 'c']
#     }
#     dummy_df = pd.DataFrame(dummy_data)
#     dummy_file = 'dummy_data.xlsx'
#     dummy_df.to_excel(dummy_file, index=False)
#
#     print(f"Created dummy file: {dummy_file}")
#     loaded_df = load_data(dummy_file, date_col='MonthYear', leading_col='IndicatorA', target_col='IndicatorB')
#
#     if loaded_df is not None:
#         print("\nLoaded DataFrame head:")
#         print(loaded_df.head())
#         print("\nDataFrame Info:")
#         loaded_df.info()
#
#     # Clean up dummy file
#     import os
#     os.remove(dummy_file)
#     print(f"\nRemoved dummy file: {dummy_file}")
