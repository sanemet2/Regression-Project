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
    parser.add_argument("--date-col", default="date",
                        help="Name of the date column in the Excel file")
    parser.add_argument("--leading-col", default="icsa",
                        help="Name of the leading indicator column")
    parser.add_argument("--target-col", default="unrate",
                        help="Name of the target indicator column")

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

    # --- Prepare Exclusion Periods --- 
    exclusion_periods_to_use = args.exclude_period
    if not exclusion_periods_to_use: # Only ask interactively if none provided via CLI
        while True:
            prompt_interactive = input("\nDo you want to specify any date periods to exclude? (y/n): ").lower()
            if prompt_interactive in ['y', 'n']:
                break
            else:
                print("Invalid input. Please enter 'y' or 'n'.")
        
        if prompt_interactive == 'y':
            interactive_exclusions = []
            print("Enter exclusion periods (format YYYY-MM-DD). Leave start date blank to finish.")
            while True:
                start_str = input("  Start date (YYYY-MM-DD) or leave blank to finish: ").strip()
                if not start_str:
                    break # Exit loop if start date is blank
                
                end_str = input(f"  End date (YYYY-MM-DD) for period starting {start_str}: ").strip()
                
                # Validate dates
                try:
                    start_date = pd.to_datetime(start_str, format='%Y-%m-%d', errors='raise')
                    end_date = pd.to_datetime(end_str, format='%Y-%m-%d', errors='raise')
                    
                    if start_date > end_date:
                        print(f"  Error: Start date {start_str} cannot be after end date {end_str}. Please re-enter.")
                        continue # Ask for the same period again
                    
                    # If valid, format and store
                    interactive_exclusions.append(f"{start_str}:{end_str}")
                    print(f"    -> Period {start_str}:{end_str} added.")
                    
                except ValueError:
                    print("  Error: Invalid date format. Please use YYYY-MM-DD. Please re-enter.")
                    continue # Ask for the same period again
            
            if interactive_exclusions:
                exclusion_periods_to_use = interactive_exclusions


    # --- Parameter Summary --- 
    print("\n--- Running Analysis With Parameters ---")
    print(f"File Path: {args.file_path}")
    print(f"Date Column: {args.date_col}")
    print(f"Leading Column: {args.leading_col}")
    print(f"Target Column: {args.target_col}")
    print(f"Sheet: {args.sheet}") 
    print(f"Header Row Index: {args.header}")
    print(f"Lead/Lag Range: -{args.range} to +{args.range}")
    print(f"Rolling Window: {args.window}")
    print(f"Output Directory: {args.output_dir}")
    if exclusion_periods_to_use: # Use the determined list
        print(f"Exclude Periods: {', '.join(exclusion_periods_to_use)}")
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
    print(f"Columns: Date='{args.date_col}', Leading='{args.leading_col}', Target='{args.target_col}'")
    print(f"Header Row Index: {args.header}")
    print(f"Output Directory: {args.output_dir}")
    if exclusion_periods_to_use: # Use the determined list
        print(f"Exclude Periods: {', '.join(exclusion_periods_to_use)}")
    print("-" * 25)

    # --- Step 2: Load Data ---
    print("Loading data...")
    # Load into df_original
    df_original = load_data(args.file_path, args.date_col, args.leading_col, args.target_col, args.header, args.sheet)

    # if df is None:
    if df_original is None:
        print("Exiting due to data loading error.")
        return # Exit if data loading failed

    # --- Step 2.5: Apply Date Exclusions (if any) ---
    # Filter the DataFrame based on user-provided exclusion periods (CLI or interactive).
    exclusion_mask = pd.Series(False, index=df_original.index) # Initialize mask for original df
    applied_exclusions = False # Flag to check if any exclusions were actually applied
    if exclusion_periods_to_use: # Check if the list (either from CLI or interactive) is not empty
        print("Parsing date exclusions...")
        # original_rows = len(df)
        original_rows = len(df_original)
        # Start with a mask where nothing is excluded
        # exclusion_mask = pd.Series(False, index=df.index) # Moved up

        for period_str in exclusion_periods_to_use: # Use the determined list
            try:
                start_str, end_str = period_str.split(':')
                # Use errors='coerce' here for robustness, already validated if interactive
                start_date = pd.to_datetime(start_str, errors='coerce') 
                end_date = pd.to_datetime(end_str, errors='coerce')

                # Check if dates parsed correctly and start <= end
                if pd.isna(start_date) or pd.isna(end_date):
                    print(f"  Warning: Could not parse dates in exclusion period '{period_str}'. Expected format YYYY-MM-DD. Skipping.")
                    continue
                if start_date > end_date:
                    print(f"  Warning: Start date {start_str} is after end date {end_str} in exclusion period '{period_str}'. Skipping this period.")
                    continue

                # Update the mask: True for rows within this period
                period_mask = (df_original.index >= start_date) & (df_original.index <= end_date)
                exclusion_mask |= period_mask # Combine masks using OR
                print(f"  Marked period {start_str} to {end_str} for exclusion.")

            except ValueError:
                print(f"  Warning: Invalid format for exclusion period '{period_str}'. Expected START:END. Skipping.")
        
        # Create df_filtered by applying the combined mask
        df_filtered = df_original[~exclusion_mask].copy() # Use ~exclusion_mask to keep non-excluded rows
        
        if len(df_filtered) < original_rows:
            applied_exclusions = True
            print(f"Data filtered. Original rows: {original_rows}, Filtered rows: {len(df_filtered)}")
        else:
            print("No rows were excluded based on the provided periods.")
            df_filtered = df_original.copy() # Ensure df_filtered exists even if no rows were excluded
            
    else:
        # If no exclusion periods provided, df_filtered is just a copy of df_original
        df_filtered = df_original.copy()
        print("No date exclusion periods specified.")

    # --- Step 3: Prepare Data for Analysis --- 
    # Use df_filtered for lead/lag analysis
    if df_filtered is None or df_filtered.empty:
        print("Error: No data available for analysis after filtering. Exiting.")
        return
    
    # --- Step 4: Find Optimal Lead/Lag --- 
    # Pass the potentially filtered dataframe here
    best_shift_filtered, r2_results_filtered = find_optimal_lead_lag(df_filtered, args.range)

    if best_shift_filtered is None:
        print("Could not determine optimal shift from filtered data. Further analysis might be unreliable.")
        # Decide how to proceed - perhaps exit, or use a default shift?
        # For now, we'll proceed but warn the user.
        best_shift = 0 # Default to 0 shift if calculation fails
    else:
        best_shift = best_shift_filtered # Use the shift found from filtered data

    # --- Step 5: Generate Plots (based on filtered R2/shift, but plotting original data) ---
    print("\n--- Plotting Optimal Shift Results (Step 5) ---")
    # Always plot using df_original, but use best_shift determined from df_filtered
    plot_scatter(df_original, best_shift, args.output_dir, args.leading_col, args.target_col)
    plot_optimal_lead(df_original, best_shift, args.output_dir, args.leading_col, args.target_col)

    # --- Step 6: Calculate Rolling Correlations --- 
    # Calculate rolling correlations on the ORIGINAL, unfiltered data
    rolling_corr_df = calculate_rolling_correlations(df_original, args.range, args.window, args.leading_col, args.target_col)

    # --- Step 6.5: Calculate Cumulative Correlations --- 
    # Calculate cumulative correlations on the ORIGINAL, unfiltered data
    cumulative_corr_df = calculate_cumulative_correlations(df_original, args.range, args.leading_col, args.target_col)
    
    # --- Step 7: Plot Rolling Correlations --- 
    # Plot rolling correlations (calculated from original data)
    plot_rolling_correlations(rolling_corr_df, args.window, args.output_dir)

    # --- Step 7.5: Plot Cumulative Correlations (NEW) --- 
    # plot_cumulative_correlations(cumulative_corr_df, args.output_dir) # Add this function call if the plotting function exists
    # Currently, the plotting function for cumulative correlation is not in plotting.py
    # For now, we skip plotting cumulative correlations, but we export the data.
    print("\nSkipping plot for cumulative correlations (functionality not implemented in plotting.py). Data will be exported.")

    # --- Step 8: Export Results to Excel --- 
    # Export using original data, but pass best_shift (from filtered), and correlation DFs (from original)
    export_to_excel(df_original, best_shift, rolling_corr_df, cumulative_corr_df, args.output_dir, args.leading_col, args.target_col, args.range, args.window)

    print("\n--- Analysis Complete --- ")
    print(f"Results exported to: {os.path.join(args.output_dir, 'analysis_results.xlsx')}")
    print(f"Plots saved in: {args.output_dir}")

if __name__ == "__main__":
    main()
