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

def export_to_excel(df_original, best_shift, rolling_corr_df, cumulative_corr_df, output_dir, leading_col_name, target_col_name, max_shift, window):
    """
    Exports the analysis results to an Excel file with multiple sheets.

    Args:
        df_original (pd.DataFrame): Original DataFrame with 'Leading', 'Target' columns and DatetimeIndex.
        best_shift (int): The optimal shift period found.
        rolling_corr_df (pd.DataFrame): DataFrame with rolling correlations per shift.
        cumulative_corr_df (pd.DataFrame): DataFrame with cumulative correlations per shift.
        output_dir (str): Directory to save the Excel file.
        leading_col_name (str): Original name of the leading column.
        target_col_name (str): Original name of the target column.
        max_shift (int): The maximum shift range used (args.range), needed for bandwidth calculation.
        window (int): Rolling correlation window size.
    """
    print(f"\n--- Exporting Results (Step 8) ---")
    # Default to 'results' if output_dir is None or empty, though argparse should handle this
    if not output_dir:
        output_dir = 'results'
    output_filename = os.path.join(output_dir, 'analysis_results.xlsx')

    try:
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)

        # Prepare data for the 'Optimal Shift Data' sheet
        optimal_df = None
        if df_original is not None and best_shift is not None:
            # Find best rolling shift
            last_rolling_corr = rolling_corr_df.iloc[-1] if not rolling_corr_df.empty else pd.Series(dtype=float)
            best_rolling_col = last_rolling_corr.idxmax()
            best_rolling_shift = int(best_rolling_col.split('_')[-1])

            # Find best cumulative shift
            last_cumulative_corr = cumulative_corr_df.iloc[-1] if not cumulative_corr_df.empty else pd.Series(dtype=float)
            best_cumulative_col = last_cumulative_corr.idxmax()
            best_cumulative_shift = int(best_cumulative_col.split('_')[-1])

            optimal_df = pd.DataFrame({
                # Keep original column names for clarity in Excel
                target_col_name: df_original['Target'],
                f'{leading_col_name}_Shifted_Roll_{window}p_{best_rolling_shift}p': df_original['Leading'].shift(best_rolling_shift),
                f'{leading_col_name}_Shifted_Cumul_{best_cumulative_shift}p': df_original['Leading'].shift(best_cumulative_shift)
            })
            optimal_df.index.name = 'Date' # Name the index column
        else:
            print("Warning: Cannot create Optimal Shift Data sheet (missing input).")

        # --- Add extra row for positive shifts ---
        # Determine the maximum positive shift between rolling and cumulative
        max_positive_shift = max(best_rolling_shift if best_rolling_shift > 0 else 0, best_cumulative_shift if best_cumulative_shift > 0 else 0)

        if optimal_df is not None and max_positive_shift > 0 and not df_original.empty:
            try:
                # Find the last valid original leading value needed based on the max shift
                # We need the value from 'max_positive_shift' periods ago to appear in the last row.
                # The value at df_original.iloc[-1] corresponds to shift 0 in the last output row.
                # The value at df_original.iloc[-max_positive_shift] will be shifted forward to the last date.
                # To fill the *next* date, we need the value from df_original.iloc[-1 - (max_positive_shift-1)]? No, simpler:
                # The value needed for the *next* row (Shift N) is the original value from the *current* last date (df_original.iloc[-1]).

                last_date = optimal_df.index[-1] # Use optimal_df index now
                # Calculate the next date based on the frequency of the index
                if pd.api.types.is_datetime64_any_dtype(optimal_df.index):
                    # Attempt to infer frequency, default to MonthBegin if fails
                    freq = pd.infer_freq(optimal_df.index)
                    if freq is None:
                        freq = pd.offsets.MonthBegin(1) # Assume monthly if cannot infer
                        print(f"  - Warning: Could not infer date frequency, assuming monthly ('{freq.name}').")
                    else:
                         print(f"  - Inferred date frequency: {freq}")
                    next_date = last_date + pd.tseries.frequencies.to_offset(freq)
                else:
                    print("  - Warning: Index is not datetime, cannot calculate next date for extra row.")
                    next_date = None

                if next_date is not None:
                    # Prepare data for the new row
                    new_row_data = {target_col_name: pd.NA}
                    if best_rolling_shift > 0:
                         roll_col_name = f'{leading_col_name}_Shifted_Roll_{window}p_{best_rolling_shift}p'
                         # Value for next date (shift N) is the *last* original value
                         new_row_data[roll_col_name] = df_original['Leading'].iloc[-1] if best_rolling_shift == 1 else df_original['Leading'].iloc[-best_rolling_shift] # Adjust source index for shift

                    if best_cumulative_shift > 0:
                        cumul_col_name = f'{leading_col_name}_Shifted_Cumul_{best_cumulative_shift}p'
                         # Value for next date (shift N) is the *last* original value
                        new_row_data[cumul_col_name] = df_original['Leading'].iloc[-1] if best_cumulative_shift == 1 else df_original['Leading'].iloc[-best_cumulative_shift] # Adjust source index for shift

                    # Add columns that might be missing in new_row_data but exist in optimal_df
                    for col in optimal_df.columns:
                         if col not in new_row_data:
                             new_row_data[col] = pd.NA

                    new_row = pd.DataFrame(new_row_data, index=[next_date])
                    new_row.index.name = 'Date' # Ensure index name matches

                    optimal_df = pd.concat([optimal_df, new_row])
                    print(f"  - Added extra row for {next_date.strftime('%Y-%m-%d')} due to positive shift.")
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
                if not r2_results_df.empty and cumulative_r2_col_name in r2_results_df:
                     # Ensure the column exists and is not all NaN before finding idxmax
                    if r2_results_df[cumulative_r2_col_name].notna().any():
                        best_cumul_r2_idx = r2_results_df[cumulative_r2_col_name].idxmax()
                        best_cumul_shift_val = r2_results_df.loc[best_cumul_r2_idx, 'Shift']
                        # Find the row index in the DataFrame where 'Shift' matches best_cumul_shift_val
                        # Add 1 for header row, and 1 because Excel is 1-based
                        target_row_excel = r2_results_df[r2_results_df['Shift'] == best_cumul_shift_val].index[0] + 2
                        # Apply bold format to that specific row
                        worksheet_r2.set_row(target_row_excel - 1, cell_format=bold_format)
                        print(f"  - Applied bold format to R2 Results row {target_row_excel} (Shift {best_cumul_shift_val} based on max Final Cumulative R2).")
                    else:
                        print("  - Warning: Cumulative R2 column is all NaN, cannot find max for bolding.")
                else:
                    print("  - Warning: Cannot find Cumulative R2 column or DataFrame is empty, skipping bold format.")
            except Exception as e:
                print(f"  - Warning: Error applying bold format to R2 Results sheet - {e}")

            # --- 2. Optimal Shift Data Sheet --- ##
            if optimal_df is not None:
                optimal_df.to_excel(writer, sheet_name='Optimal Shift Data', index=True)
                worksheet_opt = writer.sheets['Optimal Shift Data']
                # Set column widths for Optimal Shift Data sheet
                worksheet_opt.set_column(0, 0, 12) # Date column
                worksheet_opt.set_column(1, optimal_df.shape[1], 15) # Other data columns
                print("  - Optimal Shift Data sheet written.")

            # --- 3. Rolling Correlations Sheet --- ##
            if rolling_corr_df is not None and not rolling_corr_df.empty:
                rolling_corr_df.to_excel(writer, sheet_name=f'Rolling Corrs ({window}p)', index=True)
                worksheet_roll = writer.sheets[f'Rolling Corrs ({window}p)']
                _apply_correlation_formatting(worksheet_roll, rolling_corr_df, max_shift, bold_format, highlight_format, workbook, apply_bolding=True, apply_highlighting=True)
                # Set Date column width
                worksheet_roll.set_column(0, 0, 12)
                print(f"  - Wrote 'Rolling Corrs ({window}p)' sheet.")
            else:
                print("  - Skipping Rolling Corrs sheet (no data).")

            # --- 4. Cumulative Correlations Sheet --- ##
            if cumulative_corr_df is not None and not cumulative_corr_df.empty:
                cumulative_corr_df.to_excel(writer, sheet_name='Cumulative Corrs', index=True)
                worksheet_cumul = writer.sheets['Cumulative Corrs']
                _apply_correlation_formatting(worksheet_cumul, cumulative_corr_df, max_shift, bold_format, highlight_format, workbook, apply_bolding=True, apply_highlighting=True)
                # Set Date column width
                worksheet_cumul.set_column(0, 0, 12)
                print("  - Wrote 'Cumulative Corrs' sheet.")
            else:
                print("  - Skipping Cumulative Corrs sheet (no data).")

        print(f"Results successfully exported to {output_filename}")

    except Exception as e:
        print(f"Error exporting results to Excel: {e}")
