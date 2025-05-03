import argparse
import os
import pandas as pd
import sys

# Add project root to the Python path
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

from data_loader import load_data
from analysis import find_optimal_lead_lag, calculate_rolling_correlations, calculate_cumulative_correlations
from plotting import plot_scatter, plot_optimal_lead, plot_rolling_correlations
from export import export_to_excel

def main():
    parser = argparse.ArgumentParser(
        description="Time Series Lead/Lag Analysis Tool. Analyzes correlation between a leading and target indicator over time, identifying optimal lead/lag.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter # Shows defaults in help message
    )

    # Optional arguments for core data identification (with defaults)
    parser.add_argument("--file-path", default="C:\\Users\\franc\\OneDrive\\Desktop\\Programming\\Regression Project\\fredgraph.xlsx",
                        help="Path to the Excel file (.xlsx or .xls)")
    parser.add_argument("--lead-date-col", default="observation_date",
                        help="Name of the date column for the leading series")
    parser.add_argument("--lead-val-col", default="ICSA",
                        help="Name of the value column for the leading series")
    parser.add_argument("--target-date-col", default="date",
                        help="Name of the date column for the target series")
    parser.add_argument("--target-val-col", default="unrate",
                        help="Name of the value column for the target series")

    # Optional arguments for analysis parameters (will prompt if not provided)
    parser.add_argument("--range", type=int, 
                        help="Maximum lead/lag shift range (e.g., 12 periods forward/backward)")
    parser.add_argument("--window", type=int, 
                        help="Rolling window size for correlation (e.g., 36 periods)")

    # Optional arguments for file loading and output (with defaults)
    parser.add_argument("--header", type=int, default=0, 
                        help="Row number (0-indexed) containing the header")
    parser.add_argument("--sheet", default="Monthly", 
                        help="Name or index (0-indexed) of the sheet to read")
    parser.add_argument("--output_dir", default="results", 
                        help="Directory to save results")

    # --- Optional Data Exclusion ---
    parser.add_argument(
        '--exclude-period',
        action='append',  # Allow multiple instances of this argument
        type=str,
        help='Specify a period to exclude (format: YYYY-MM-DD:YYYY-MM-DD). Can be used multiple times.',
        metavar='START_DATE:END_DATE',
        default=[] # Initialize as empty list if not provided
    )

    args = parser.parse_args()

    # --- Interactive Prompts for Optional Parameters --- 
    if args.range is None:
        while True:
            try:
                prompt = (
                    "Please enter the maximum lead/lag shift range (e.g., 12).\n" 
                    "This defines how many periods forward and backward (+/- N) to test for the best correlation: "
                )
                range_input = input(prompt)
                args.range = int(range_input)
                if args.range > 0:
                    break
                else:
                    print("Error: Range must be a positive integer.")
            except ValueError:
                print("Error: Invalid input. Please enter an integer.")

    if args.window is None:
        while True:
            try:
                prompt = (
                    "Please enter the rolling window size (e.g., 36).\n" 
                    "This is the number of periods used to calculate the rolling correlation: "
                )
                window_input = input(prompt)
                args.window = int(window_input)
                if args.window > 0:
                    break
                else:
                    print("Error: Window size must be a positive integer.")
            except ValueError:
                print("Error: Invalid input. Please enter an integer.")

    # --- Parameter Summary --- 
    print("\n--- Running Analysis With Parameters ---")
    print(f"File Path: {args.file_path}")
    print(f"Leading Date Column: {args.lead_date_col}")
    print(f"Leading Value Column: {args.lead_val_col}")
    print(f"Target Date Column: {args.target_date_col}")
    print(f"Target Value Column: {args.target_val_col}")
    print(f"Sheet: {args.sheet}") 
    print(f"Header Row Index: {args.header}")
    print(f"Lead/Lag Range: -{args.range} to +{args.range}")
    print(f"Rolling Window: {args.window}")
    print(f"Output Directory: {args.output_dir}")
    if args.exclude_period: # Use the determined list
        print(f"Exclude Periods: {', '.join(args.exclude_period)}")
    print("-" * 38)

    # Create output directory if it doesn't exist
    if not os.path.exists(args.output_dir):
        try:
            os.makedirs(args.output_dir)
            print(f"Created output directory: {args.output_dir}")
        except OSError as e:
            print(f"Error creating output directory {args.output_dir}: {e}")
            return

    print("--- Starting Analysis ---")
    print(f"File: {args.file_path}")
    print(f"Sheet: {args.sheet}") # Print sheet name/index
    print(f"Leading Date Column: {args.lead_date_col}, Leading Value Column: {args.lead_val_col}")
    print(f"Target Date Column: {args.target_date_col}, Target Value Column: {args.target_val_col}")
    print(f"Header Row Index: {args.header}")
    print(f"Output Directory: {args.output_dir}")
    if args.exclude_period: # Use the determined list
        print(f"Exclude Periods: {', '.join(args.exclude_period)}")
    print("-" * 25)

    # --- Step 2: Load Data ---
    print("\n--- Loading Data (Step 2) ---")
    lead_series_raw, target_series_raw = load_data(
        file_path=args.file_path,
        lead_date_col=args.lead_date_col,
        lead_val_col=args.lead_val_col,
        target_date_col=args.target_date_col,
        target_val_col=args.target_val_col,
        header_row=args.header,
        sheet_name=args.sheet
    )

    if lead_series_raw is None or target_series_raw is None:
        print("Exiting - Data loading failed or one/both series could not be created.")
        return # Exit if data loading failed

    # --- Step 3: Detect Individual Frequencies ---
    print("\n--- Detecting Frequencies (Step 3) ---")
    freq_lead = None
    if isinstance(lead_series_raw.index, pd.DatetimeIndex):
        freq_lead = pd.infer_freq(lead_series_raw.index)
        print(f"  Detected frequency for Leading Series ('{args.lead_val_col}' based on '{args.lead_date_col}'): {freq_lead or 'Irregular'}")
    else:
        # This case should ideally not happen if load_data succeeds
        print(f"  Warning: Leading series index is not DatetimeIndex.")

    freq_target = None
    if isinstance(target_series_raw.index, pd.DatetimeIndex):
        freq_target = pd.infer_freq(target_series_raw.index)
        print(f"  Detected frequency for Target Series ('{args.target_val_col}' based on '{args.target_date_col}'): {freq_target or 'Irregular'}")
    else:
        # This case should ideally not happen if load_data succeeds
        print(f"  Warning: Target series index is not DatetimeIndex.")

    # --- Step 4: Check Resampling Condition ---
    print("\n--- Checking Resampling Condition (Step 4) ---")
    resampling_needed = False
    if freq_lead and freq_target and (freq_lead != freq_target):
        resampling_needed = True
        print(f"  Resampling needed: Frequencies differ ('{freq_lead}' vs '{freq_target}').")
    elif not freq_lead or not freq_target:
        print("  Resampling not possible: Could not determine a regular frequency for one or both series.")
    else:
        # Frequencies are the same
        print(f"  Resampling not needed: Frequencies match ('{freq_lead}').")

    # --- Step 5 & 6: Determine Target Frequency, Series and Resample (if needed) ---
    lead_series_final = lead_series_raw
    target_series_final = target_series_raw

    if resampling_needed:
        print("\n--- Preparing for Resampling (Steps 5 & 6) ---")
        try:
            # --- Explicit Weekly vs Monthly Handling ---
            is_lead_weekly = freq_lead and freq_lead.upper().startswith('W')
            is_target_weekly = freq_target and freq_target.upper().startswith('W')
            is_lead_monthly = freq_lead and freq_lead.upper().startswith('M')
            is_target_monthly = freq_target and freq_target.upper().startswith('M')

            resampled = False
            if is_lead_weekly and is_target_monthly:
                target_frequency = 'MS' # Target is MS
                print(f"  Detected Weekly ('{freq_lead}') vs Monthly ('{freq_target}'). Resampling Leading to '{target_frequency}'.")
                lead_series_final = lead_series_raw.resample(target_frequency).last()
                lead_series_final = lead_series_final.rename('Leading')
                target_series_final = target_series_raw # Keep monthly target
                print(f"  Resampled Leading Series to {target_frequency} using '.last()'. New length: {len(lead_series_final)}")
                resampled = True
            elif is_target_weekly and is_lead_monthly:
                target_frequency = 'MS' # Target is MS
                print(f"  Detected Monthly ('{freq_lead}') vs Weekly ('{freq_target}'). Resampling Target to '{target_frequency}'.")
                target_series_final = target_series_raw.resample(target_frequency).last()
                target_series_final = target_series_final.rename('Target')
                lead_series_final = lead_series_raw # Keep monthly lead
                print(f"  Resampled Target Series to {target_frequency} using '.last()'. New length: {len(target_series_final)}")
                resampled = True

            # --- Fallback to Offset Comparison (if not handled above) ---
            if not resampled:
                print("  Attempting frequency comparison using offsets...")
                offset_lead = pd.tseries.frequencies.to_offset(freq_lead)
                offset_target = pd.tseries.frequencies.to_offset(freq_target)

                if offset_lead.nanos > offset_target.nanos: # Lead is lower frequency
                    target_frequency = freq_lead
                    series_to_resample = target_series_raw
                    series_to_keep = lead_series_raw
                    print(f"  Target frequency (lower): {target_frequency} (from Leading)")
                    print(f"  Series to resample (higher): Target ('{args.target_val_col}') from {freq_target}")

                    target_series_final = series_to_resample.resample(target_frequency).last()
                    target_series_final = target_series_final.rename('Target')
                    lead_series_final = series_to_keep
                    print(f"  Resampled Target Series to {target_frequency} using '.last()'. New length: {len(target_series_final)}")
                    resampled = True

                elif offset_target.nanos > offset_lead.nanos: # Target is lower frequency
                    target_frequency = freq_target
                    series_to_resample = lead_series_raw
                    series_to_keep = target_series_raw
                    print(f"  Target frequency (lower): {target_frequency} (from Target)")
                    print(f"  Series to resample (higher): Leading ('{args.lead_val_col}') from {freq_lead}")

                    lead_series_final = series_to_resample.resample(target_frequency).last()
                    lead_series_final = lead_series_final.rename('Leading')
                    target_series_final = series_to_keep
                    print(f"  Resampled Leading Series to {target_frequency} using '.last()'. New length: {len(lead_series_final)}")
                    resampled = True
                else:
                    print(f"  Warning: Frequencies '{freq_lead}' and '{freq_target}' are different strings but resolve to the same offset or comparison failed. No resampling performed.")
                    # Keep original series if comparison doesn't yield a difference
                    lead_series_final = lead_series_raw
                    target_series_final = target_series_raw

        except ValueError as e:
             # Error during offset conversion or comparison
             print(f"  Error during frequency comparison/conversion '{freq_lead}', '{freq_target}': {e}. Cannot resample.")
             # Keep original series if comparison fails
             lead_series_final = lead_series_raw
             target_series_final = target_series_raw
        except Exception as e:
             # Catch any other unexpected errors during resampling
             print(f"  An unexpected error occurred during resampling preparation: {e}")
             lead_series_final = lead_series_raw
             target_series_final = target_series_raw
    else:
        # If resampling not needed, the final series are just the raw ones (already assigned)
        pass

    # --- Step 7: Combine and Align Data ---
    print("\n--- Combining and Aligning Data (Step 7) ---")
    # Create the final analysis DataFrame by combining the (potentially resampled) series.
    # Pandas automatically aligns on the index.
    df_analysis = pd.DataFrame({
        'Leading': lead_series_final,
        'Target': target_series_final
    })

    # Drop rows where *either* column has NaN after alignment/resampling
    initial_rows = len(df_analysis)
    df_analysis.dropna(inplace=True)
    dropped_rows = initial_rows - len(df_analysis)

    if dropped_rows > 0:
        print(f"  Dropped {dropped_rows} rows due to missing values after alignment.")

    if df_analysis.empty:
        print("Error: No overlapping data remains after aligning the two series. Cannot proceed.")
        return
    else:
        print(f"  Final aligned dataset 'df_analysis' created. Shape: {df_analysis.shape}")
        print(f"  Date range: {df_analysis.index.min()} to {df_analysis.index.max()}")

    # --- Analysis Steps (Now using df_analysis) ---

    # --- Step 4: Find Optimal Lead/Lag (using analysis.py function) ---
    optimal_shift, r2_results_df = find_optimal_lead_lag(
        df_analysis, # Use the combined & aligned DataFrame
        max_shift=args.range
        # No need to specify lead/target cols, function uses 'Leading'/'Target'
    )
    if optimal_shift is None:
        print("Exiting - Optimal lead/lag calculation failed.")
        return

    # --- Step 5: Prepare Shifted Data (using optimal shift) ---
    print("\n--- Preparing Optimally Shifted Data (Step 5) ---")
    df_shifted = df_analysis.copy()
    df_shifted['Leading_Shifted'] = df_shifted['Leading'].shift(optimal_shift)
    df_analysis_final = df_shifted[['Leading_Shifted', 'Target']].dropna()
    print(f"Created final df with optimally shifted leading series. Shape: {df_analysis_final.shape}")

    # --- Step 6: Calculate Rolling Correlation (using analysis.py function) ---
    rolling_corr_df = calculate_rolling_correlations(
        df_analysis, # Use the combined & aligned (pre-shift) DataFrame for rolling calc
        max_shift=args.range,
        window=args.window
        # No need to specify lead/target cols, function uses 'Leading'/'Target'
    )
    if rolling_corr_df is None:
        print("Warning: Rolling correlation calculation failed or returned None.")

    # --- Calculate Cumulative Correlations (using analysis.py function) ---
    cumulative_corr_df = calculate_cumulative_correlations(
        df_analysis, # Use the combined & aligned (pre-shift) DataFrame
        max_shift=args.range
        # No need to specify lead/target cols, function uses 'Leading'/'Target'
    )
    if cumulative_corr_df is None:
        print("Warning: Cumulative correlation calculation failed or returned None.")

    # --- Step 7: Export Results (using export.py function) ---
    print("\n--- Exporting Results (Step 7) ---")
    print("Exporting results to Excel...")
    export_to_excel(df_aligned_data=df_analysis, # Pass the aligned DataFrame
                          best_shift=optimal_shift,
                          rolling_corr_df=rolling_corr_df,
                          cumulative_corr_df=cumulative_corr_df,
                          output_dir=args.output_dir,
                          max_shift=args.range,
                          window=args.window)

    print("\nAnalysis complete.")


if __name__ == "__main__":
    main()
