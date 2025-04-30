import pandas as pd
import numpy as np

def find_optimal_lead_lag(df, max_shift):
    """
    Calculates R-squared for different lead/lag shifts of the 'Leading' series
    against the 'Target' series and finds the optimal shift.

    Args:
        df (pd.DataFrame): DataFrame with 'Leading' and 'Target' columns,
                           and a DatetimeIndex.
        max_shift (int): The maximum number of periods to shift (e.g., 12 for -12 to +12).

    Returns:
        tuple: (best_shift (int), r2_results (pd.DataFrame))
               best_shift: The lead/lag period with the highest R-squared.
               r2_results: DataFrame with columns ['Shift', 'R_Squared'].
               Returns (None, None) if calculation fails.
    """
    results = []
    target_series = df['Target']
    leading_series = df['Leading']

    shifts_to_test = range(-max_shift, max_shift + 1)

    print(f"\n--- Lead/Lag Analysis (Step 4) ---")
    print(f"Testing shifts from {min(shifts_to_test)} to {max(shifts_to_test)}...")

    for shift in shifts_to_test:
        shifted_leading = leading_series.shift(shift)
        temp_df = pd.DataFrame({'Target': target_series, 'Shifted_Leading': shifted_leading}).dropna()

        if len(temp_df) < 2:
            r_squared = np.nan
        else:
            correlation = temp_df['Shifted_Leading'].corr(temp_df['Target'])
            r_squared = correlation ** 2

        results.append({'Shift': shift, 'R_Squared': r_squared})

    r2_results_df = pd.DataFrame(results)

    if r2_results_df['R_Squared'].isnull().all():
        print("Warning: Could not calculate R-squared for any shift. Not enough overlapping data.")
        return None, r2_results_df

    best_result = r2_results_df.loc[r2_results_df['R_Squared'].idxmax()]
    best_shift = int(best_result['Shift'])

    print(f"Optimal Shift Found: {best_shift} periods (R-Squared: {best_result['R_Squared']:.4f})")

    return best_shift, r2_results_df

# --- Rolling Correlation Function ---
def calculate_rolling_correlations(df, max_shift, window):
    """
    Calculates rolling correlations for different lead/lag shifts.

    Args:
        df (pd.DataFrame): DataFrame with 'Leading' and 'Target' columns, DatetimeIndex.
        max_shift (int): The maximum number of periods to shift (-max_shift to +max_shift).
        window (int): The rolling window size for the correlation calculation.

    Returns:
        pd.DataFrame: A DataFrame where index is date, columns are shift periods,
                      and values are the rolling correlations. Returns None if error.
    """
    print(f"\n--- Rolling Correlation (Step 6) ---")
    print(f"Calculating {window}-period rolling correlations for shifts {-max_shift} to {max_shift}...")

    if window <= 1:
        print("Error: Rolling window must be greater than 1.")
        return None
    if df.empty:
        print("Error: Input DataFrame is empty.")
        return None

    target_series = df['Target']
    leading_series = df['Leading']
    rolling_corr_results = {} # Dictionary to store series for each shift

    shifts_to_test = range(-max_shift, max_shift + 1)

    for shift in shifts_to_test:
        shifted_leading = leading_series.shift(shift)

        # Calculate rolling correlation between target and shifted leading series
        min_periods_required = int(window * 0.9) # Require at least 90% of window to have data
        rolling_corr = shifted_leading.rolling(window=window, min_periods=min_periods_required).corr(target_series)

        # Store the resulting series, naming it by the shift
        rolling_corr_results[f'Shift_{shift}'] = rolling_corr

        # Optional progress print
        # if shift % 5 == 0:
        #     print(f"  Calculated rolling correlation for shift {shift}")

    # Combine all resulting series into a single DataFrame
    try:
        rolling_corr_df = pd.DataFrame(rolling_corr_results)
        print(f"Rolling correlations calculated. Shape: {rolling_corr_df.shape}")
        # Optional: Print head/tail to check
        # print(rolling_corr_df.dropna().head())
    except Exception as e:
        print(f"Error combining rolling correlation results into DataFrame: {e}")
        return None

    return rolling_corr_df
