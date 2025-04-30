import argparse
import os
from data_loader import load_data
from analysis import find_optimal_lead_lag, calculate_rolling_correlations
from plotting import plot_scatter, plot_optimal_lead, plot_rolling_correlations
from export import export_to_excel

def main():
    parser = argparse.ArgumentParser(description="Time Series Lead/Lag Analysis Tool")

    # --- Input Arguments ---
    parser.add_argument("file_path", help="Path to the input Excel file (.xlsx or .xls)")
    parser.add_argument("date_col", help="Name of the date column in the Excel file")
    parser.add_argument("leading_col", help="Name of the leading indicator column")
    parser.add_argument("target_col", help="Name of the target indicator column")
    parser.add_argument("--header", type=int, default=0,
                        help="0-indexed row number containing column headers (default: 0)")
    parser.add_argument("--sheet", default=0, 
                        help="Name or 0-indexed position of the Excel sheet to read (default: 0)")

    # --- Analysis Parameters ---
    parser.add_argument("-r", "--range", type=int, required=True,
                        help="Maximum number of periods to shift for lead/lag analysis (e.g., 12 for -12 to +12)")
    parser.add_argument("-w", "--window", type=int, required=True,
                        help="Rolling window size for correlation calculation (e.g., 36)")

    # --- Output Arguments ---
    parser.add_argument("-o", "--output_dir", default="results",
                        help="Directory to save output plots and Excel file (default: results)")

    args = parser.parse_args()

    # --- Validate Inputs ---
    if not os.path.exists(args.file_path):
        print(f"Error: Input file not found at {args.file_path}")
        return

    if args.range <= 0:
        print("Error: Lead/lag range must be a positive integer.")
        return

    if args.window <= 0:
         print("Error: Rolling window size must be a positive integer.")
         return

    if args.header < 0:
         print("Error: Header row index must be 0 or greater.")
         return

    # Convert sheet argument if it's numeric (for index)
    sheet_arg = args.sheet
    try:
        sheet_arg = int(sheet_arg) # Try converting to int for index
    except ValueError:
        pass # Keep as string if it's not purely numeric

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
    print(f"Sheet: {sheet_arg}") # Print sheet name/index
    print(f"Columns: Date='{args.date_col}', Leading='{args.leading_col}', Target='{args.target_col}'")
    print(f"Header Row Index: {args.header}")
    print(f"Lead/Lag Range: -{args.range} to +{args.range}")
    print(f"Rolling Window: {args.window}")
    print(f"Output Directory: {args.output_dir}")
    print("-" * 25)

    # --- Step 2: Load Data ---
    df = load_data(args.file_path, args.date_col, args.leading_col, args.target_col, args.header, sheet_arg)

    if df is None:
        print("Exiting due to data loading errors.")
        return # Stop execution if data loading failed

    print("\n--- Data Loading Complete ---")
    print(df.head()) # Print head to confirm loading

    # --- Step 4: Lead/Lag Analysis ---
    best_shift, r2_results = find_optimal_lead_lag(df, args.range)

    # Check if analysis was successful
    if best_shift is None:
        print("Lead/lag analysis could not determine an optimal shift. Exiting.")
        # Optionally print the r2_results DataFrame here for debugging
        if r2_results is not None:
            print("\nR-Squared results per shift:")
            print(r2_results.to_string()) # Use to_string to avoid truncation
        return

    # Optional: Print the R-squared results DataFrame
    print("\nR-Squared results per shift:")
    print(r2_results.to_string()) # Use to_string() to show all rows if needed

    # --- Plotting (Step 5) ---
    print("\n--- Plotting (Step 5) ---")
    plot_scatter(df, best_shift, args.output_dir, args.leading_col, args.target_col)
    plot_optimal_lead(df, best_shift, args.output_dir, args.leading_col, args.target_col)

    # --- Rolling Correlation (Step 6) ---
    rolling_corr_df = calculate_rolling_correlations(df, args.range, args.window)

    # --- Plotting Rolling Correlation (Step 7) ---
    plot_rolling_correlations(rolling_corr_df, args.window, args.output_dir)

    # --- Exporting Results (Step 8) ---
    export_to_excel(r2_results, df, best_shift, rolling_corr_df,
                    args.output_dir, args.leading_col, args.target_col)

    print("\n--- Analysis Complete ---")


if __name__ == "__main__":
    main()
