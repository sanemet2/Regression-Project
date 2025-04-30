import pandas as pd

def load_data(file_path, date_col, leading_col, target_col, header_row=0, sheet_name=0):
    """
    Loads data from a specific sheet in an Excel file, selects specified columns,
    parses dates, and handles missing values.

    Args:
        file_path (str): Path to the Excel file.
        date_col (str): Name of the date column.
        leading_col (str): Name of the leading series column.
        target_col (str): Name of the target series column.
        header_row (int): The 0-indexed row number containing the headers.
        sheet_name (str or int): Name or 0-indexed position of the sheet to read.

    Returns:
        pandas.DataFrame: Processed DataFrame with selected columns,
                          or None if an error occurs.
    """
    try:
        # Determine the engine based on file extension
        if file_path.endswith('.xlsx'):
            engine = 'openpyxl'
        elif file_path.endswith('.xls'):
            engine = 'xlrd' # Note: xlrd might be needed for .xls
        else:
            print(f"Error: Unsupported file format for {file_path}. Please use .xlsx or .xls.")
            return None

        df = pd.read_excel(file_path, engine=engine, header=header_row, sheet_name=sheet_name)
        print(f"Info: Reading sheet '{sheet_name}' with headers from row index {header_row}.")

        required_cols = [date_col, leading_col, target_col]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            print(f"Error: The following columns were not found in sheet '{sheet_name}' (using header row {header_row}): {', '.join(missing_cols)}")
            print(f"Available columns: {', '.join(df.columns)}")
            return None

        df = df[required_cols].copy()

        df.rename(columns={
            date_col: 'Date',
            leading_col: 'Leading',
            target_col: 'Target'
        }, inplace=True)

        try:
            df['Date'] = pd.to_datetime(df['Date'])
        except ValueError:
            try:
                df['Date'] = pd.to_datetime(df['Date'], format='%m/%y')
                print("Info: Interpreted date column using 'mm/yy' format.")
            except ValueError:
                try:
                    df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d')
                    print("Info: Interpreted date column using 'YYYY-MM-DD' format.")
                except ValueError:
                    print(f"Error: Could not parse the date column '{date_col}'. Please ensure it's in a recognizable format (e.g., YYYY-MM-DD, MM/DD/YYYY, mm/yy).")
                    return None
        except Exception as e:
            print(f"Error parsing date column '{date_col}': {e}")
            return None

        df['Leading'] = pd.to_numeric(df['Leading'], errors='coerce')
        df['Target'] = pd.to_numeric(df['Target'], errors='coerce')

        initial_rows = len(df)
        df.dropna(subset=['Date', 'Leading', 'Target'], inplace=True)
        dropped_rows = initial_rows - len(df)
        if dropped_rows > 0:
            print(f"Info: Dropped {dropped_rows} row(s) due to missing values in Date, Leading, or Target columns.")

        df.sort_values(by='Date', inplace=True)
        df.set_index('Date', inplace=True) 

        if df.empty:
            print("Error: No valid data remaining after processing and removing missing values.")
            return None

        print(f"Data loaded successfully from sheet '{sheet_name}'. Index range: {df.index.min()} to {df.index.max()}. Shape: {df.shape}")
        return df

    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
        return None
    except ValueError as e: 
         if "Worksheet" in str(e) and "does not exist" in str(e):
              print(f"Error: Sheet name '{sheet_name}' not found in the Excel file.")
              # You might want to list available sheets here if needed
              # excel_file = pd.ExcelFile(file_path, engine=engine)
              # print(f"Available sheets: {excel_file.sheet_names}")
         else:
              print(f"An unexpected error occurred during data loading: {e}")
         return None
    except Exception as e:
        print(f"An unexpected error occurred during data loading: {e}")
        return None
