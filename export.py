import pandas as pd
import os

def export_to_excel(r2_results_df, df, best_shift, rolling_corr_df, output_dir, leading_col_name, target_col_name):
    """
    Exports the analysis results to an Excel file with multiple sheets.

    Args:
        r2_results_df (pd.DataFrame): DataFrame with R-squared results per shift.
        df (pd.DataFrame): Original DataFrame with 'Leading', 'Target' columns and DatetimeIndex.
        best_shift (int): The optimal shift period found.
        rolling_corr_df (pd.DataFrame): DataFrame with rolling correlations per shift.
        output_dir (str): Directory to save the Excel file.
        leading_col_name (str): Original name of the leading column.
        target_col_name (str): Original name of the target column.
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
        if df is not None and best_shift is not None:
            shifted_leading = df['Leading'].shift(best_shift)
            optimal_df = pd.DataFrame({
                # Keep original column names for clarity in Excel
                target_col_name: df['Target'],
                f'{leading_col_name}_Shifted_{best_shift}p': shifted_leading
            })
            optimal_df.index.name = 'Date' # Name the index column
        else:
            print("Warning: Cannot create Optimal Shift Data sheet (missing input).")

        # --- Add extra row for positive shifts ---
        if optimal_df is not None and best_shift is not None and best_shift > 0 and not df.empty:
            try:
                last_date = df.index[-1]
                next_date = last_date + pd.offsets.MonthBegin(1)
                last_leading_value = df['Leading'].iloc[-1]
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
            # Write R-squared results
            if r2_results_df is not None:
                r2_results_df.to_excel(writer, sheet_name='R2 Results', index=False)
                print(f"  - R2 Results sheet written.")
            else:
                print("  - Skipping R2 Results sheet (no data).")
                pd.DataFrame([{'Status': 'R-Squared results unavailable'}]).to_excel(writer, sheet_name='R2 Results', index=False)

            # Write Optimal Shift Data
            if optimal_df is not None:
                optimal_df.to_excel(writer, sheet_name='Optimal Shift Data')
                print(f"  - Optimal Shift Data sheet written.")
            else:
                 print("  - Skipping Optimal Shift Data sheet (no data).")
                 pd.DataFrame([{'Status': 'Optimal shift data unavailable'}]).to_excel(writer, sheet_name='Optimal Shift Data', index=False)

            # Write Rolling Correlations
            if rolling_corr_df is not None:
                rolling_corr_df.to_excel(writer, sheet_name='Rolling Correlations')
                print(f"  - Rolling Correlations sheet written.")
            else:
                 print("  - Skipping Rolling Correlations sheet (no data).")
                 pd.DataFrame([{'Status': 'Rolling correlations unavailable'}]).to_excel(writer, sheet_name='Rolling Correlations', index=False)

        print(f"Results successfully exported to {output_filename}")

    except Exception as e:
        print(f"Error exporting results to Excel: {e}")
