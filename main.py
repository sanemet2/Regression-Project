import argparse
import os
import pandas as pd
import sys

# Add project root to the Python path
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

from data_loader import load_data
from analysis import find_optimal_lead_lag, calculate_rolling_correlations
from plotting import plot_scatter, plot_optimal_lead, plot_rolling_correlations
from export import export_to_excel

def main():
    parser = argparse.ArgumentParser(description="Time Series Lead/Lag Analysis Tool")

    # Required arguments for core data identification
    parser.add_argument("file_path", help="Path to the Excel file (.xlsx or .xls)")
    parser.add_argument("date_col", help="Name of the date column in the Excel file")
    parser.add_argument("leading_col", help="Name of the leading indicator column")
    parser.add_argument("target_col", help="Name of the target indicator column")

    # Optional arguments for analysis parameters (will prompt if not provided)
    parser.add_argument("--range", type=int, help="Maximum lead/lag shift range (e.g., 12)")
    parser.add_argument("--window", type=int, help="Rolling window size for correlation (e.g., 36)")

    # Optional arguments for file loading and output
    parser.add_argument("--header", type=int, default=0, help="Row number (0-indexed) containing the header. Default: 0")
    parser.add_argument("--sheet", default=0, help="Name or index (0-indexed) of the sheet to read. Default: 0")
    parser.add_argument("--output_dir", default="results", help="Directory to save results. Default: 'results'")

    args = parser.parse_args()

    # --- Interactive Prompts for Optional Parameters --- 
    if args.range is None:
        while True:
            try:
                range_input = input(f"Please enter the maximum lead/lag shift range (e.g., 12): ")
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
                window_input = input(f"Please enter the rolling window size (e.g., 36): ")
                args.window = int(window_input)
                if args.window > 0:
                    break
                else:
                    print("Error: Window size must be a positive integer.")
            except ValueError:
                print("Error: Invalid input. Please enter an integer.")
    # --- End Interactive Prompts ---

    # Create output directory if it doesn't exist
    if not os.path.exists(args.output_dir):
        try:
            os.makedirs(args.output_dir)
            print(f"Created output directory: {args.output_dir}")
        except OSError as e:
            print(f"Error creating output directory {args.output_dir}: {e}")
            sys.exit(1)

    # --- Load Data ---
    print("Loading data...")
    try:
        df = load_data(args.file_path, args.date_col, args.leading_col, args.target_col, args.header, args.sheet)
        print(f"Data loaded successfully. Shape: {df.shape}")
        # Basic validation after loading
        if df.empty:
            print("Error: Dataframe is empty after loading. Check file path, sheet name, and column names.")
            sys.exit(1)
        if df['Date'].isnull().any():
             print(f"Warning: Original date column '{args.date_col}' contains NaNs before processing.")
        if df[['Leading', 'Target']].isnull().values.any():
            print(f"Warning: Leading ('{args.leading_col}') or Target ('{args.target_col}') columns contain NaNs before processing. Rows with NaNs will be dropped.")

    except FileNotFoundError:
        print(f"Error: File not found at {args.file_path}")
        sys.exit(1)
    except KeyError as e:
        print(f"Error: Column not found in Excel file - {e}. Check column names and header row.")
        sys.exit(1)
    except ValueError as e:
        print(f"Error loading data: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred during data loading: {e}")
        sys.exit(1)

    # --- Perform Analysis ---
    print("Finding optimal lead/lag...")
    try:
        r2_results_df, best_shift, max_r2 = find_optimal_lead_lag(df.copy(), args.range)
        if r2_results_df.empty:
             print("Warning: Optimal lead/lag analysis returned empty results. Check data and range.")
        else:
             print(f"Optimal shift: {best_shift} period(s) with R-squared: {max_r2:.4f}")
    except Exception as e:
        print(f"An error occurred during optimal lead/lag analysis: {e}")
        # Optionally continue without this result or exit
        r2_results_df, best_shift, max_r2 = pd.DataFrame(), None, None 
        # sys.exit(1) # Uncomment to exit on error

    print("Calculating rolling correlations...")
    try:
        rolling_corr_df = calculate_rolling_correlations(df.copy(), args.range, args.window)
        if rolling_corr_df.empty:
             print("Warning: Rolling correlation analysis returned empty results. Check data, range and window.")
    except Exception as e:
        print(f"An error occurred during rolling correlation analysis: {e}")
        rolling_corr_df = pd.DataFrame() # Assign empty df to allow export to continue
        # sys.exit(1) # Uncomment to exit on error

    # --- Generate Plots ---
    print("Generating plots...")
    optimal_scatter_path = os.path.join(args.output_dir, "optimal_scatter.png")
    optimal_line_path = os.path.join(args.output_dir, "optimal_line.png")
    rolling_corr_path = os.path.join(args.output_dir, "rolling_correlations.png")

    try:
        if best_shift is not None:
            plot_scatter(df.copy(), best_shift, optimal_scatter_path, args.leading_col, args.target_col)
            plot_optimal_lead(df.copy(), best_shift, optimal_line_path, args.leading_col, args.target_col)
        else:
            print("Skipping scatter and optimal lead plots due to missing best_shift.")
    except Exception as e:
        print(f"Error generating optimal shift plots: {e}")

    try:
        if not rolling_corr_df.empty:
            plot_rolling_correlations(rolling_corr_df, rolling_corr_path, args.window)
        else:
             print("Skipping rolling correlation plot due to empty results.")
    except Exception as e:
        print(f"Error generating rolling correlation plot: {e}")

    # --- Export Results --- 
    print("Exporting results to Excel...")
    excel_output_path = os.path.join(args.output_dir, "analysis_results.xlsx")
    try:
        export_to_excel(r2_results_df, df.copy(), best_shift, rolling_corr_df, excel_output_path, args.leading_col, args.target_col)
        print(f"Results exported successfully to {excel_output_path}")
    except Exception as e:
        print(f"Error exporting results to Excel: {e}")

    print("\nAnalysis Complete.")

if __name__ == "__main__":
    main()