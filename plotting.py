# plotting.py
import matplotlib.pyplot as plt
import pandas as pd
import os

def plot_scatter(df, best_shift, output_dir, leading_col_name, target_col_name):
    """
    Generates and saves a scatter plot of the target series vs. the optimally
    shifted leading series.

    Args:
        df (pd.DataFrame): DataFrame with 'Leading' and 'Target' columns and DatetimeIndex.
        best_shift (int): The optimal shift period found.
        output_dir (str): Directory to save the plot.
        leading_col_name (str): Original name of the leading column for labeling.
        target_col_name (str): Original name of the target column for labeling.
    """
    try:
        shifted_leading = df['Leading'].shift(best_shift)
        temp_df = pd.DataFrame({
            'Target': df['Target'],
            f'Shifted_Leading_{best_shift}p': shifted_leading
        }).dropna() # Drop NaNs for scatter plot alignment

        if temp_df.empty:
            print("Warning: Cannot generate scatter plot. No overlapping data after shift.")
            return

        plt.figure(figsize=(10, 6))
        plt.scatter(temp_df[f'Shifted_Leading_{best_shift}p'], temp_df['Target'], alpha=0.5)
        plt.title(f'Scatter Plot: {target_col_name} vs. {leading_col_name} (Shifted {best_shift} Periods)')
        plt.xlabel(f'{leading_col_name} (Shifted {best_shift} Periods)')
        plt.ylabel(target_col_name)
        plt.grid(True)

        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        plot_filename = os.path.join(output_dir, f'scatter_plot_shift_{best_shift}.png')
        plt.savefig(plot_filename)
        plt.close() # Close the plot to free memory
        print(f"Scatter plot saved to {plot_filename}")

    except Exception as e:
        print(f"Error generating scatter plot: {e}")

def plot_optimal_lead(df, best_shift, output_dir, leading_col_name, target_col_name):
    """
    Generates and saves a line plot showing the target series and the
    optimally shifted leading series over time.

    Args:
        df (pd.DataFrame): DataFrame with 'Leading' and 'Target' columns and DatetimeIndex.
        best_shift (int): The optimal shift period found.
        output_dir (str): Directory to save the plot.
        leading_col_name (str): Original name of the leading column for labeling.
        target_col_name (str): Original name of the target column for labeling.
    """
    try:
        shifted_leading = df['Leading'].shift(best_shift)

        plt.figure(figsize=(12, 6))
        plt.plot(df.index, df['Target'], label=f'{target_col_name} (Target)')
        plt.plot(df.index, shifted_leading, label=f'{leading_col_name} (Shifted {best_shift} Periods)', alpha=0.7)

        plt.title(f'Time Series: {target_col_name} vs. Optimally Shifted {leading_col_name}')
        plt.xlabel('Date')
        plt.ylabel('Value')
        plt.legend()
        plt.grid(True)

        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        plot_filename = os.path.join(output_dir, f'line_plot_optimal_shift_{best_shift}.png')
        plt.savefig(plot_filename)
        plt.close() # Close the plot to free memory
        print(f"Optimal lead line chart saved to {plot_filename}")

    except Exception as e:
        print(f"Error generating optimal lead line chart: {e}")


# --- Rolling Correlation Plot ---
def plot_rolling_correlations(rolling_corr_df, window, output_dir): 
    """
    Generates and saves a line plot showing the evolution of rolling correlations
    for all tested shifts over time.

    Args:
        rolling_corr_df (pd.DataFrame): DataFrame with rolling correlations for each shift.
                                         Index is date, columns are like 'Shift_N'.
        window (int): The rolling window size used (for plot title).
        output_dir (str): Directory to save the plot.
    """
    print("\n--- Plotting Rolling Correlation (Step 7) ---")
    if rolling_corr_df is None or rolling_corr_df.empty:
        print("Skipping rolling correlation plot (no data).")
        return

    try:
        plt.figure(figsize=(14, 7))

        num_shifts = len(rolling_corr_df.columns)
        for col in rolling_corr_df.columns:
            plt.plot(rolling_corr_df.index, rolling_corr_df[col], label=col, alpha=0.6)

        plt.title(f'{window}-Period Rolling Correlations Over Time') 
        plt.xlabel('Date')
        plt.ylabel('Rolling Correlation')
        plt.grid(True)

        # Always display legend, even if crowded
        plt.legend(title='Shift Period', bbox_to_anchor=(1.05, 1), loc='upper left')

        os.makedirs(output_dir, exist_ok=True)
        plot_filename = os.path.join(output_dir, 'rolling_correlations_evolution.png')
        # Adjust layout to prevent legend overlap
        plt.tight_layout(rect=[0, 0, 0.85, 1]) 
        plt.savefig(plot_filename)
        plt.close()
        print(f"Rolling correlation evolution plot saved to {plot_filename}")

    except Exception as e:
        print(f"Error generating rolling correlation plot: {e}")
