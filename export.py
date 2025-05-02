import pandas as pd
import os
import math
import numpy as np

# Helper function to apply formatting to correlation sheets
def _apply_correlation_formatting(worksheet, df, max_shift, bold_format, highlight_format, workbook):
    """Applies bold to strongest correlation column and highlight to bandwidth."""
    if df is None or df.empty:
        print("  - Skipping correlation formatting (DataFrame is None or empty).")
        return

    # Check if DataFrame has at least one data row
    if len(df) == 0:
        print("  - Skipping correlation formatting (DataFrame has no data rows).")
        return

    try:
        # Get the last row of data
        last_row = df.iloc[-1]

        # Filter out non-numeric columns before finding max absolute value
        numeric_cols = df.select_dtypes(include=np.number).columns
        if not numeric_cols.any():
            print("  - Skipping correlation formatting (No numeric columns found).")
            return
            
        numeric_last_row = last_row[numeric_cols]
        if numeric_last_row.isnull().all(): # Handle if last row is all NaN
            print("  - Skipping correlation formatting (Last row is all NaN).")
            return

        # Find the column name with the max absolute value in the last row
        strongest_shift_col_name = numeric_last_row.abs().idxmax()

        # Extract the shift number 'S' from the column name (e.g., 'RollCorr_Shift_5')
        try:
            strongest_shift_S = int(strongest_shift_col_name.split('_')[-1])
        except (ValueError, IndexError):
            print(f"  - Warning: Could not parse shift number from column '{strongest_shift_col_name}'. Skipping bandwidth formatting.")
            strongest_shift_S = None # Cannot determine bandwidth without S

        # Find the Excel column index for this shift (1-based)
        try:
            col_idx_df = df.columns.get_loc(strongest_shift_col_name)
            col_idx_excel = col_idx_df + 1 # +1 for the index/Date column
        except KeyError:
             print(f"  - Warning: Strongest correlation column '{strongest_shift_col_name}' not found. Skipping formatting.")
             return

        # Apply bold format to the strongest shift column (data cells + header)
        # Adjust width slightly for bolded column
        worksheet.set_column(col_idx_excel, col_idx_excel, width=15, cell_format=bold_format)
        print(f"  - Applied bold format to column {col_idx_excel} ('{strongest_shift_col_name}').")

        # --- Apply Bandwidth Highlighting (only if strongest_shift_S was found) ---
        if strongest_shift_S is not None:
            total_shifts = 2 * max_shift + 1
            # Ensure bandwidth is at least 1 if total_shifts > 0
            bandwidth_size = max(1, round(0.25 * total_shifts)) if total_shifts > 0 else 0
            radius = math.floor(bandwidth_size / 2)
            print(f"  - Calculated bandwidth radius: {radius} (total shifts: {total_shifts}, bandwidth size: {bandwidth_size})")

            start_shift = strongest_shift_S - radius
            end_shift = strongest_shift_S + radius
            col_prefix = strongest_shift_col_name.split('_Shift_')[0] # Get prefix like 'RollCorr' or 'CumCorr'

            for shift_val in range(start_shift, end_shift + 1):
                target_col_name = f"{col_prefix}_Shift_{shift_val}"
                if target_col_name in df.columns:
                    try:
                        target_col_idx_df = df.columns.get_loc(target_col_name)
                        target_col_idx_excel = target_col_idx_df + 1

                        # Apply conditional format to highlight the column data
                        # Range: row 1 (first data row) to len(df) (last data row)
                        # Column: target_col_idx_excel
                        worksheet.conditional_format(1, target_col_idx_excel, len(df), target_col_idx_excel,
                                                 {'type': 'no_blanks', # Apply to all non-blank cells in the column
                                                  'format': highlight_format})
                        # print(f"  - Applied highlight format condition to column {target_col_idx_excel} (Shift {shift_val}).") # Too verbose
                    except KeyError:
                         pass # Should not happen if target_col_name in df.columns
                # else: # Ignore shifts in bandwidth that don't exist as columns
                #    pass 
            print(f"  - Applied highlight formatting condition for shifts {start_shift} to {end_shift}.")
        else:
             print("  - Skipping bandwidth highlighting as strongest shift number could not be parsed.")

        # Optional: Adjust general column widths for readability
        # worksheet.set_column(1, len(df.columns), 12) # Set width for all correlation columns

    except Exception as e:
        # Print specific error details
        import traceback
        print(f"  - Error applying correlation formatting: {e}")
        # traceback.print_exc() # Uncomment for detailed debugging if needed

# Helper function to extract shift from column name like 'Shift_X' or 'CumCorr_Shift_X'
def _get_shift_from_col(col_name):
    try:
        # Handle potential negative signs
        if '_-' in col_name:
             # Split by '_-' and take the last part, negate it
            return -int(col_name.split('_-')[-1])
        else:
            # Split by '_' and take the last part
            return int(col_name.split('_')[-1])
    except (IndexError, ValueError):
        # Fallback or error handling if pattern doesn't match
        print(f"Warning: Could not extract shift number from column name '{col_name}'. Defaulting to 0.")
        return 0

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
            shifted_leading = df_original['Leading'].shift(best_shift)
            # Find best rolling shift
            last_rolling_corr = rolling_corr_df.iloc[-1].abs()
            best_rolling_col = last_rolling_corr.idxmax()
            best_rolling_shift = _get_shift_from_col(best_rolling_col)
            shifted_roll_leading = df_original['Leading'].shift(best_rolling_shift)

            # Find best cumulative shift
            last_cumulative_corr = cumulative_corr_df.iloc[-1].abs()
            best_cumulative_col = last_cumulative_corr.idxmax()
            best_cumulative_shift = _get_shift_from_col(best_cumulative_col)
            shifted_cumul_leading = df_original['Leading'].shift(best_cumulative_shift)

            optimal_df = pd.DataFrame({
                # Keep original column names for clarity in Excel
                target_col_name: df_original['Target'],
                f'{leading_col_name}_Shifted_{best_shift}p': shifted_leading,
                f'{leading_col_name}_Shifted_Roll_{window}p_{best_rolling_shift}p': shifted_roll_leading,
                f'{leading_col_name}_Shifted_Cumul_{best_cumulative_shift}p': shifted_cumul_leading
            })
            optimal_df.index.name = 'Date' # Name the index column
        else:
            print("Warning: Cannot create Optimal Shift Data sheet (missing input).")

        # --- Add extra row for positive shifts ---
        if optimal_df is not None and best_shift is not None and best_shift > 0 and not df_original.empty:
            try:
                last_date = df_original.index[-1]
                next_date = last_date + pd.offsets.MonthBegin(1)
                last_leading_value = df_original['Leading'].iloc[-1]
                shifted_col_name = f'{leading_col_name}_Shifted_{best_shift}p'

                new_row_data = {target_col_name: pd.NA, shifted_col_name: last_leading_value}
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
            bold_format = workbook.add_format({'bold': True})
            highlight_format = workbook.add_format({'bg_color': '#E0E0E0'}) # Light grey background

            # --- 1. R2 Results Sheet --- ##
            # Prepare data for R2 Results sheet enhancement
            final_rolling_corr = rolling_corr_df.iloc[-1] if not rolling_corr_df.empty else pd.Series(dtype=float)
            final_cumulative_corr = cumulative_corr_df.iloc[-1] if not cumulative_corr_df.empty else pd.Series(dtype=float)

            # Create maps from shift number to final correlation value
            final_roll_corr_map = { _get_shift_from_col(col): val for col, val in final_rolling_corr.items() }
            final_cumul_corr_map = { _get_shift_from_col(col): val for col, val in final_cumulative_corr.items() }

            # Determine the set of all shifts tested from the maps
            all_shifts = sorted(list(set(final_roll_corr_map.keys()) | set(final_cumul_corr_map.keys())))

            # Create the final R2 results DataFrame structure
            r2_results_list = []
            for shift in all_shifts:
                final_roll = final_roll_corr_map.get(shift, np.nan)
                final_cumul = final_cumul_corr_map.get(shift, np.nan)
                r2_results_list.append({
                    'Shift': shift,
                    f'R2 (Final Rolling - {window}p)': final_roll**2 if not pd.isna(final_roll) else np.nan,
                    'R2 (Final Cumulative)': final_cumul**2 if not pd.isna(final_cumul) else np.nan,
                    f'Final Rolling Corr ({window}p)': final_roll,
                    'Final Cumulative Corr': final_cumul
                })
            
            r2_results_final_df = pd.DataFrame(r2_results_list)
            # Set 'Shift' as the index for easier lookup and cleaner sheet
            r2_results_final_df.set_index('Shift', inplace=True)
            
            # Sort by index (Shift)
            r2_results_final_df.sort_index(inplace=True)

            # Write R2 Results sheet
            if not r2_results_final_df.empty:
                r2_results_final_df.to_excel(writer, sheet_name='R2 Results', index=True) # Write index (Shift)
                worksheet_r2 = writer.sheets['R2 Results']
                
                # Find the shift (index) with the max 'R2 (Final Cumulative)'
                try:
                    shift_to_bold = r2_results_final_df['R2 (Final Cumulative)'].idxmax()
                    # Find the row number in Excel (0-indexed header + index of the shift in the sorted index)
                    row_to_bold = r2_results_final_df.index.get_loc(shift_to_bold) + 1 # +1 for header
                    # Apply bold format to the entire row (column A to last column)
                    last_col_r2 = len(r2_results_final_df.columns) # Number of columns (index is not counted here)
                    worksheet_r2.set_row(row_to_bold, cell_format=bold_format)
                    print(f"  - Applied bold format to row {row_to_bold + 1} in 'R2 Results' (Shift {shift_to_bold}).")
                except ValueError: # Handle case where all R2 values are NaN
                    print("  - Skipping bold formatting in 'R2 Results' (all R2 values are NaN or sheet empty).")
                except Exception as e:
                    print(f"  - Error applying bold formatting to 'R2 Results' sheet: {e}")

                # Auto-adjust column widths for R2 Results
                for i, col in enumerate(r2_results_final_df.reset_index().columns): # Use reset_index to include 'Shift' in width calc
                    width = max(len(str(col)), r2_results_final_df.reset_index()[col].astype(str).map(len).max()) + 1
                    worksheet_r2.set_column(i, i, width)
                print("  - Auto-adjusted column widths for 'R2 Results'.")
            else:
                print("  - Skipping 'R2 Results' sheet (no data).")

            # --- 2. Optimal Shift Data Sheet ---
            if optimal_df is not None and not optimal_df.empty:
                optimal_df.to_excel(writer, sheet_name='Optimal Shift Data', index=True) # Write index (Date)
                worksheet_opt = writer.sheets['Optimal Shift Data']
                # Auto-adjust column widths
                for i, col in enumerate(optimal_df.reset_index().columns):
                    width = max(len(str(col)), optimal_df.reset_index()[col].astype(str).map(len).max()) + 2 # Add padding
                    worksheet_opt.set_column(i, i, width)
                print("  - Wrote 'Optimal Shift Data' sheet.")
            else:
                print("  - Skipping 'Optimal Shift Data' sheet (no data).")

            # --- 3. Rolling Correlations Sheet ---
            if rolling_corr_df is not None and not rolling_corr_df.empty:
                rolling_corr_df.to_excel(writer, sheet_name=f'Rolling Corrs ({window}p)', index=True)
                worksheet_roll = writer.sheets[f'Rolling Corrs ({window}p)']
                _apply_correlation_formatting(worksheet_roll, rolling_corr_df, max_shift, bold_format, highlight_format, workbook)
                # Set Date column width
                worksheet_roll.set_column(0, 0, 12) 
                print(f"  - Wrote 'Rolling Corrs ({window}p)' sheet.")
            else:
                print(f"  - Skipping 'Rolling Corrs ({window}p)' sheet (no data).")

            # --- 4. Cumulative Correlations Sheet ---
            if cumulative_corr_df is not None and not cumulative_corr_df.empty:
                cumulative_corr_df.to_excel(writer, sheet_name='Cumulative Corrs', index=True)
                worksheet_cumul = writer.sheets['Cumulative Corrs']
                _apply_correlation_formatting(worksheet_cumul, cumulative_corr_df, max_shift, bold_format, highlight_format, workbook)
                # Set Date column width
                worksheet_cumul.set_column(0, 0, 12) 
                print("  - Wrote 'Cumulative Corrs' sheet.")
            else:
                print("  - Skipping 'Cumulative Corrs' sheet (no data).")

        print(f"Results successfully exported to {output_filename}")

    except Exception as e:
        print(f"Error exporting results to Excel: {e}")
        import traceback
        traceback.print_exc()
