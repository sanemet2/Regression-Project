import pandas as pd
import os
import math
import numpy as np

# Helper function to apply formatting to correlation sheets
def _apply_correlation_formatting(worksheet, df, max_shift, bold_format, highlight_format, workbook, apply_bolding=True, apply_highlighting=True):
    """Applies conditional formatting to a correlation DataFrame in Excel."""
    if df.empty:
        return

    # Determine the strongest shift based on the absolute value of the last row's correlation
    # Ensure we only consider numeric columns for idxmax
    numeric_cols = df.select_dtypes(include=np.number)
    if numeric_cols.empty:
        print("  - Warning: No numeric columns found to determine strongest shift. Skipping formatting.")
        return
    last_row_corr = numeric_cols.iloc[-1].abs()
    if last_row_corr.empty:
        print("  - Warning: Last row is empty or non-numeric. Skipping formatting.")
        return
    strongest_shift_col_name = last_row_corr.idxmax()

    # Extract the base prefix (e.g., 'Shift' or 'CumCorr') and the shift value
    col_prefix = ""
    strongest_shift_S = None
    try:
        parts = strongest_shift_col_name.split('_Shift_')
        if len(parts) == 2:
            # Case: 'Prefix_Shift_X'
            col_prefix = parts[0]
            strongest_shift_S = int(parts[1])
        elif len(parts) == 1 and strongest_shift_col_name.startswith('Shift_'):
            # Case: 'Shift_X'
            col_prefix = "Shift" # Use a standard prefix internally
            strongest_shift_S = int(strongest_shift_col_name.split('_')[-1])
        else:
             raise ValueError("Unrecognized column name format")
    except (ValueError, IndexError):
        print(f"  - Warning: Could not parse prefix/shift from column '{strongest_shift_col_name}'. Skipping formatting.")
        return # Cannot proceed without prefix and shift

    # Find the Excel column index for this shift (1-based)
    try:
        col_idx_df = df.columns.get_loc(strongest_shift_col_name)
        col_idx_excel = col_idx_df + 1 # +1 because Excel columns are 1-indexed
    except KeyError:
        print(f"  - Warning: Strongest shift column '{strongest_shift_col_name}' not found in DataFrame index. Skipping formatting.")
        return

    max_row = len(df) # Number of data rows

    # --- Apply Bolding --- ##
    if apply_bolding:
        # Use conditional format to apply bold to data cells
        worksheet.conditional_format(1, col_idx_excel, max_row, col_idx_excel,
                                     {'type': 'no_blanks', # Apply to all non-blank cells
                                      'format': bold_format})
        # Optionally make header bold too (apply format directly)
        # worksheet.write(0, col_idx_excel, df.columns[col_idx_df], bold_format) # Requires getting header string
        # Adjust width slightly for bolded column
        worksheet.set_column(col_idx_excel, col_idx_excel, width=12)
        print(f"  - Applied bold format condition to column {col_idx_excel} ('{strongest_shift_col_name}').")
    else:
        # Ensure default width if not bolding
        worksheet.set_column(col_idx_excel, col_idx_excel, width=12) # Adjust if needed

    # --- Apply Bandwidth Highlighting --- ##
    if apply_highlighting:
        total_shifts = 2 * max_shift + 1
        bandwidth_size = max(1, round(0.25 * total_shifts)) if total_shifts > 0 else 0
        if bandwidth_size % 2 == 0: # Ensure bandwidth is odd for symmetry
            bandwidth_size += 1
        radius = (bandwidth_size - 1) // 2
        start_shift = strongest_shift_S - radius
        end_shift = strongest_shift_S + radius

        col_idx_excel_start = None
        col_idx_excel_end = None
        for shift_val in range(start_shift, end_shift + 1):
            # *** FIX: Construct name using the *correct* prefix ***
            # Handle cases where prefix might be empty (e.g. for 'Shift_X') or present ('CumCorr_Shift_X')
            if col_prefix == "Shift": # Special handling for the base 'Shift_X' case
                target_col_name = f"Shift_{shift_val}"
            else:
                target_col_name = f"{col_prefix}_Shift_{shift_val}"
            if target_col_name in df.columns:
                try:
                    target_col_idx_df = df.columns.get_loc(target_col_name)
                    target_col_idx_excel = target_col_idx_df + 1
                    if col_idx_excel_start is None:
                        col_idx_excel_start = target_col_idx_excel
                    col_idx_excel_end = target_col_idx_excel # Always update end index
                except KeyError:
                    pass
            #else:
                #print(f"      DEBUG: Column {target_col_name} not found in DataFrame.")

        if col_idx_excel_start is not None and col_idx_excel_end is not None:
            worksheet.conditional_format(1, col_idx_excel_start, max_row, col_idx_excel_end,
                                         {'type': 'no_blanks', 'format': highlight_format})
            print(f"  - Applied highlight formatting condition for shifts {start_shift} to {end_shift}.")
        else:
            print(f"  - Warning: Could not find columns for entire highlight range {start_shift} to {end_shift}. Skipping highlight.")

def export_to_excel(df_aligned_data, best_shift, rolling_corr_df, cumulative_corr_df, output_dir, max_shift, window):
    """
    Exports the analysis results to an Excel file with multiple sheets.

    Args:
        df_aligned_data (pd.DataFrame): DataFrame with 'Leading', 'Target' columns and DatetimeIndex, aligned.
        best_shift (int): The optimal shift period found (currently based on R2, might need revisit).
        rolling_corr_df (pd.DataFrame): DataFrame with rolling correlations per shift.
        cumulative_corr_df (pd.DataFrame): DataFrame with cumulative correlations per shift.
        output_dir (str): Directory to save the Excel file.
        max_shift (int): The maximum shift range used (args.range), needed for bandwidth calculation.
        window (int): Rolling correlation window size.
    """
    print(f"\n--- Exporting Results (Step 8) ---")
    # Default to 'results' if output_dir is None or empty
    if not output_dir:
        output_dir = 'results'
    output_filename = os.path.join(output_dir, 'analysis_results.xlsx')

    try:
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)

        # Prepare data for the 'Optimal Shift Data' sheet
        optimal_df = None
        if df_aligned_data is not None and best_shift is not None:
            # Find best rolling shift based on LAST value
            best_rolling_shift = None
            if rolling_corr_df is not None and not rolling_corr_df.empty:
                last_rolling_corr = rolling_corr_df.iloc[-1].abs()
                best_rolling_col = last_rolling_corr.idxmax() if not last_rolling_corr.empty else None
                if best_rolling_col:
                    try:
                        best_rolling_shift = int(best_rolling_col.split('_')[-1])
                    except (ValueError, IndexError):
                        print(f"Warning: Could not parse shift from rolling column '{best_rolling_col}'")

            # Find best cumulative shift based on LAST value
            best_cumulative_shift = None
            if cumulative_corr_df is not None and not cumulative_corr_df.empty:
                last_cumulative_corr = cumulative_corr_df.iloc[-1].abs()
                best_cumulative_col = last_cumulative_corr.idxmax() if not last_cumulative_corr.empty else None
                if best_cumulative_col:
                     try:
                         best_cumulative_shift = int(best_cumulative_col.split('_')[-1])
                     except (ValueError, IndexError):
                         print(f"Warning: Could not parse shift from cumulative column '{best_cumulative_col}'")

            # Check if we actually found shifts before proceeding
            if best_rolling_shift is not None and best_cumulative_shift is not None:
                optimal_df = pd.DataFrame({
                    # Use fixed 'Target' and 'Leading' names internally
                    'Target': df_aligned_data['Target'],
                    f'Leading_Shifted_Roll_{window}p_{best_rolling_shift}p': df_aligned_data['Leading'].shift(best_rolling_shift),
                    f'Leading_Shifted_Cumul_{best_cumulative_shift}p': df_aligned_data['Leading'].shift(best_cumulative_shift)
                })
                optimal_df.index.name = 'Date' # Name the index column
            else:
                print("Warning: Could not determine best rolling or cumulative shift from last values. Optimal Shift Data sheet might be incomplete.")
                # Create a basic df if one shift was found, or empty if none
                optimal_df = pd.DataFrame({'Target': df_aligned_data['Target']})
                if best_rolling_shift is not None:
                     optimal_df[f'Leading_Shifted_Roll_{window}p_{best_rolling_shift}p'] = df_aligned_data['Leading'].shift(best_rolling_shift)
                if best_cumulative_shift is not None:
                     optimal_df[f'Leading_Shifted_Cumul_{best_cumulative_shift}p'] = df_aligned_data['Leading'].shift(best_cumulative_shift)
                optimal_df.index.name = 'Date'

        else:
            print("Warning: Cannot create Optimal Shift Data sheet (missing aligned data or best_shift).")

        # --- Add extra row for positive shifts (Based on R2 best_shift for now) ---
        # TODO: Decide if this should use rolling/cumulative best shift?
        if optimal_df is not None and best_shift is not None and best_shift > 0 and df_aligned_data is not None and not df_aligned_data.empty:
            try:
                last_date = df_aligned_data.index[-1]
                # Determine the frequency of the aligned data for correct date offset
                freq = pd.infer_freq(df_aligned_data.index)
                if freq:
                    offset = pd.tseries.frequencies.to_offset(freq)
                    next_date = last_date + offset
                else:
                     # Attempt a guess if frequency is irregular (e.g., assume monthly)
                    print("Warning: Could not infer frequency for aligned data index. Assuming monthly offset for extra row.")
                    next_date = last_date + pd.offsets.MonthBegin(1)

                last_leading_value = df_aligned_data['Leading'].iloc[-1]

                # Need to figure out which shifted column to add to, if any exist.
                # Let's just add a generic 'Leading_Shifted_Future' for now if needed
                # Or maybe add NA to all existing shifted columns?
                new_row_data = {'Target': pd.NA}
                for col in optimal_df.columns:
                    if col.startswith('Leading_Shifted'):
                        new_row_data[col] = last_leading_value
                    elif col != 'Target': # Add NA to any other unexpected columns
                        new_row_data[col] = pd.NA

                # If no shifted columns exist yet (e.g., if above logic failed),
                # create a placeholder based on the original 'best_shift' (from R2)
                if not any(col.startswith('Leading_Shifted') for col in optimal_df.columns):
                    shifted_col_name = f'Leading_Shifted_R2_{best_shift}p' # Use R2 shift
                    new_row_data[shifted_col_name] = last_leading_value

                new_row = pd.DataFrame(new_row_data, index=[next_date])
                new_row.index.name = 'Date' # Ensure index name matches

                optimal_df = pd.concat([optimal_df, new_row])
                print(f"  - Added extra row for {next_date.strftime('%Y-%m-%d')} due to positive shift (using last value of 'Leading').")
            except Exception as e:
                print(f"Warning: Could not add extra shifted row - {e}")
        # --- End extra row addition ---

        with pd.ExcelWriter(output_filename, engine='xlsxwriter', datetime_format='yyyy-mm-dd') as writer:
            # Get workbook and define formats
            workbook = writer.book
            bold_format = workbook.add_format({'bold': True}) # Restore bold format
            highlight_format = workbook.add_format({'bg_color': '#E0E0E0'}) # Restore light grey background

            # --- 1. R2 Results Sheet --- ##
            # Prepare data for R2 Results sheet enhancement
            final_rolling_corr = rolling_corr_df.iloc[-1] if not rolling_corr_df.empty else pd.Series(dtype=float)
            final_cumulative_corr = cumulative_corr_df.iloc[-1] if not cumulative_corr_df.empty else pd.Series(dtype=float)

            # Create maps from shift number to final correlation value
            final_roll_corr_map = { int(col.split('_')[-1]): val for col, val in final_rolling_corr.items() }
            final_cumul_corr_map = { int(col.split('_')[-1]): val for col, val in final_cumulative_corr.items() }

            # Determine the set of all shifts tested from the maps
            all_shifts = sorted(list(set(final_roll_corr_map.keys()) | set(final_cumul_corr_map.keys())))

            # Create initial DataFrame with just shifts
            r2_results_df = pd.DataFrame({'Shift': all_shifts})

            # Add new R2 columns using the maps, squaring the correlation values
            rolling_r2_col_name = f'R2 (Final Rolling - {window}p)'
            cumulative_r2_col_name = 'R2 (Final Cumulative)'
            r2_results_df[rolling_r2_col_name] = r2_results_df['Shift'].map(final_roll_corr_map).apply(lambda x: x**2 if pd.notna(x) else np.nan)
            r2_results_df[cumulative_r2_col_name] = r2_results_df['Shift'].map(final_cumul_corr_map).apply(lambda x: x**2 if pd.notna(x) else np.nan)

            # Write the enhanced DataFrame to Excel
            r2_results_df.to_excel(writer, sheet_name='R2 Results', index=False)
            worksheet_r2 = writer.sheets['R2 Results']

            # Find the row number for the shift with the highest FINAL CUMULATIVE R2
            try:
                # Ensure the column exists and has data before finding max
                if cumulative_r2_col_name in r2_results_df and r2_results_df[cumulative_r2_col_name].notna().any():
                    best_cumulative_r2_idx = r2_results_df[cumulative_r2_col_name].idxmax()
                    best_shift_for_bolding = r2_results_df.loc[best_cumulative_r2_idx, 'Shift']
                    best_shift_excel_row = best_cumulative_r2_idx + 1 # Excel rows are 1-based, +1 for header
                    worksheet_r2.set_row(best_shift_excel_row, None, bold_format)
                    print(f"  - Applied bold format to R2 Results row {best_shift_excel_row + 1} (Shift {best_shift_for_bolding} based on max Final Cumulative R2).")
                else:
                    print(f"  Warning: Cannot determine row to bold. '{cumulative_r2_col_name}' column missing or empty.")
            except Exception as e: # Catch potential errors during max finding or indexing
                print(f"  Warning: Could not apply bold format to R2 Results sheet. Error: {e}")

            # Autofit columns for R2 Results
            for i, col in enumerate(r2_results_df.columns):
                # Calculate max width needed for column header and data
                col_data = r2_results_df[col].astype(str)
                max_len = max(len(str(col)), col_data.map(len).max()) + 2 # Add buffer
                worksheet_r2.set_column(i, i, max_len)

            # --- 2. Optimal Shift Data Sheet --- ##
            if optimal_df is not None:
                optimal_df.to_excel(writer, sheet_name='Optimal Shift Data')
                print(f"  - Optimal Shift Data sheet written.")
            else:
                 print("  - Skipping Optimal Shift Data sheet (no data).")
                 pd.DataFrame([{'Status': 'Optimal shift data unavailable'}]).to_excel(writer, sheet_name='Optimal Shift Data', index=False)

            # --- 3. Rolling Correlations Sheet ---
            if rolling_corr_df is not None and not rolling_corr_df.empty:
                rolling_corr_df.to_excel(writer, sheet_name=f'Rolling Corrs ({window}p)', index=True)
                worksheet_roll = writer.sheets[f'Rolling Corrs ({window}p)']
                _apply_correlation_formatting(worksheet_roll, rolling_corr_df, max_shift, bold_format, highlight_format, workbook, apply_bolding=True, apply_highlighting=True)
                # Set Date column width
                worksheet_roll.set_column(0, 0, 12) 
                print(f"  - Wrote 'Rolling Corrs ({window}p)' sheet.")
            else:
                print("  - Skipping Rolling Correlations sheet (no data).")
                pd.DataFrame([{'Status': 'Rolling correlations unavailable'}]).to_excel(writer, sheet_name=f'Rolling Corrs ({window}p)', index=False)

            # --- 4. Cumulative Correlations Sheet ---
            if cumulative_corr_df is not None and not cumulative_corr_df.empty:
                cumulative_corr_df.to_excel(writer, sheet_name='Cumulative Corrs', index=True)
                worksheet_cumul = writer.sheets['Cumulative Corrs']
                _apply_correlation_formatting(worksheet_cumul, cumulative_corr_df, max_shift, bold_format, highlight_format, workbook, apply_bolding=True, apply_highlighting=True)
                # Set Date column width
                worksheet_cumul.set_column(0, 0, 12) 
                print("  - Wrote 'Cumulative Corrs' sheet.")
            else:
                print("  - Skipping Cumulative Correlations sheet (no data).")
                pd.DataFrame([{'Status': 'Cumulative correlations unavailable'}]).to_excel(writer, sheet_name='Cumulative Corrs', index=False)

            # --- 5. Original Aligned Data Sheet --- ##
            if df_aligned_data is not None and not df_aligned_data.empty:
                df_aligned_data.to_excel(writer, sheet_name='Aligned Input Data')
                worksheet_data = writer.sheets['Aligned Input Data']
                worksheet_data.set_column(0, 0, 12) # Date column width
                for i, col in enumerate(df_aligned_data.columns):
                    width = max(len(col), df_aligned_data[col].apply(lambda x: len(str(x)) if pd.notna(x) else 0).max()) + 2
                    worksheet_data.set_column(i + 1, i + 1, min(width, 50))
            else:
                print("  - Skipping 'Aligned Input Data' sheet (no data).")

        print(f"Results successfully exported to {output_filename}")

    except Exception as e:
        print(f"Error exporting results to Excel: {e}")
