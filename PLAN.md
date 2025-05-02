# Project Plan: Time Series Lead/Lag Analysis App

## 1. Goal

Create a Python Command Line Interface (CLI) application that analyzes the lead/lag relationship between two time series from an Excel file. The app will identify the optimal lead/lag time based on R-squared, visualize the relationship, calculate rolling correlations for different leads/lags, and export results to Excel.

## 2. Core Features (Implemented in CLI)

*   **Data Loading:**
    *   Loads data from a user-specified Excel file (`.xlsx` or `.xls`).
    *   Allows the user to specify:
        *   Date column name.
        *   Leading series column name.
        *   Target series column name.
    *   Parses dates, attempting common formats including `mm/yy`.
    *   Handles missing data by dropping rows with `NaN` in selected columns.
    *   Optionally excludes user-specified date periods (e.g., `--exclude-period START:END`) before analysis.
*   **Optimal Lead/Lag Analysis:**
    *   User specifies the maximum lead/lag period `N`.
    *   The application tests shifts from `-N` to `+N` periods.
    *   For each shift, it aligns the series and calculates the R-squared value between the shifted leading series and the target series.
    *   Identifies the shift (lead/lag) with the highest R-squared.
    *   Outputs:
        *   Prints the optimal lead/lag period and the corresponding R-squared value.
        *   Generates and saves a scatter plot of the optimally shifted leading series vs. the target series.
        *   Generates and saves a line chart showing the target series and the optimally shifted leading series.
*   **Rolling Correlation Analysis:**
    *   User specifies the rolling window size.
    *   For each lead/lag tested (`-N` to `+N`), calculates the rolling correlation between the shifted leading series and the target series using the specified window.
    *   Outputs:
        *   Generates and saves a line chart plotting the rolling correlations for *all* tested leads/lags over time.
*   **Cumulative Correlation Analysis:**
    *   For each lead/lag tested (`-N` to `+N`), calculates the cumulative correlation between the shifted leading series and the target series.
    *   Outputs:
        *   Generates and saves a line chart plotting the cumulative correlations for *all* tested leads/lags over time.
*   **Export Results:**
    *   Creates an output Excel file (`analysis_results.xlsx`).
    *   Includes sheets for:
        *   `R2 Results`: Table showing each tested lead/lag and its R-squared value.
        *   `Optimal Shift Data`: Columns for Date, Target Series, Leading Series (optimally shifted).
        *   `Rolling Correlations`: Columns for Date and the rolling correlation values for each tested lead/lag.
        *   `Cumulative Correlations`: Columns for Date and the cumulative correlation values for each tested lead/lag.
    *   Saves generated plots as separate image files.

## 3. Input Handling Strategy (CLI)

*   **File Path & Core Columns (`--file-path`, `--date-col`, `--leading-col`, `--target-col`):** These are optional command-line arguments. Default values (`C:\Users\franc\OneDrive\Desktop\Programming\Regression Project\fredgraph.xlsx`, `date`, `icsa`, `unrate`, respectively) are hardcoded in `main.py` and shown in the help message (`-h`). The script uses these defaults if the arguments are not provided and does not prompt interactively for them.

*   **Analysis Parameters (`--range`, `--window`):** These are optional command-line arguments. If not provided via the command line, the script prompts the user interactively with descriptions of what the parameters mean. Basic positive integer validation is performed on the interactive input.

*   **Other Parameters (`--header`, `--sheet`, `--output_dir`):** These remain optional command-line arguments with defaults (`0`, `Monthly`, `results`, respectively). The script does not prompt interactively for these.

*   **Date Exclusion (`--exclude-period`):** Optionally specify date periods to remove from the analysis *before* calculations. Use the format `YYYY-MM-DD:YYYY-MM-DD`. This argument can be used multiple times to exclude several distinct periods.

## 4. Technical Stack

*   **Language:** Python 3.x
*   **Core Libraries:**
    *   `pandas`: Data loading, manipulation, date handling, rolling calculations.
    *   `numpy`: Numerical operations.
    *   `scikit-learn`: Calculating R-squared (`sklearn.metrics.r2_score`).
    *   `matplotlib` / `seaborn`: Generating plots.
    *   `openpyxl` / `xlsxwriter`: Reading/writing Excel files.
    *   `argparse`: Handling command-line arguments.

## 5. Future Considerations

*   GUI Development (e.g., using Tkinter, PyQt, or Streamlit).
*   Embedding plots directly into Excel (using `xlsxwriter` capabilities).
*   Automatically adjust time periods if one is weekly, the other monthly etc.
*   Fix charting functions and rolling correlation charts
*   Fix formatting


## 6. Output Formatting & Exclusion Logic Refinements

*   **Goal:** Improve the readability of the Excel output with targeted formatting and adjust the application of date exclusions to specific analysis steps.

*   **Implementation Steps:**

    *   [x] **Refactor Data Handling for Exclusions:**
        *   Modify `main.py` to load data into an `df_original`.
        *   If exclusions are specified, create `df_filtered`. Otherwise, `df_filtered = df_original`.
        *   Pass `df_filtered` ONLY to `find_optimal_lead_lag` and `plot_scatter`.
        *   Pass `df_original` to correlation calculations, other plots, and `export_to_excel`.
        *   *(Status: Done)*

    *   [x] **Update `export_to_excel` Function (`export.py`):**
        *   Pass `max_shift` parameter for formatting calculations.
        *   *(Status: Done)*

    *   [x] **Style 'R2 Results' Sheet:**
        *   In `export_to_excel`, find the row corresponding to `best_shift` (derived from filtered data if exclusions applied).
        *   Apply a bold format to that entire row.
        *   *(Status: Done)*

    *   [x] **Style 'Rolling Correlations' Sheet:**
        *   Add helper `_apply_correlation_formatting`.
        *   Bold the column with the highest absolute correlation in the *last* row (using `df_original` data).
        *   Calculate bandwidth (e.g., 25% of total shifts).
        *   Apply conditional highlighting (e.g., light background) to columns within the bandwidth around the strongest correlation.
        *   *(Status: Done)*

    *   [x] **Style 'Cumulative Correlations' Sheet:**
        *   Apply the same formatting logic as 'Rolling Correlations' sheet (using `df_original` data).
        *   *(Status: Done)*

    *   [x] **Enhance 'Optimal Shift Data' Sheet:**
        *   Determine `best_rolling_shift` and `best_cumulative_shift` from the last row of respective correlation DataFrames (using `df_original` data).
        *   Include target column and leading column shifted by `best_shift` (R2), `best_rolling_shift`, and `best_cumulative_shift`.
        *   Use descriptive column names.
        *   *(Status: Done)*

    *   [x] **Enhance 'R2 Results' Sheet (Summary):**
        *   Modify `export_to_excel` in `export.py`.
        *   Extract final rolling correlation value for each shift from the last row of `rolling_corr_df`.
        *   Extract final cumulative correlation value for each shift from the last row of `cumulative_corr_df`.
        *   Calculate the R-squared (`r^2`) for these final correlation values.
        *   Remove the original `R_Squared` column (redundant with `R2 (Final Cumulative)`).
        *   Add 'R2 (Final Rolling - {window}p)' and 'R2 (Final Cumulative)' columns to the `r2_results_df` before writing to Excel.
        *   Update bold formatting to highlight the row with the maximum value in the `R2 (Final Cumulative)` column.
        *   *(Status: Done)*

    *   [x] **Add Window Period to Excel Headers:**
        *   Modify `export_to_excel` in `export.py`.
        *   Include the rolling window size (e.g., `{window}p`) in relevant column headers.
            *   'R2 Results': Change `R2 (Final Rolling)` to `R2 (Final Rolling - {window}p)`.
            *   'Optimal Shift Data': Change the shifted rolling correlation column name to include window and shift (e.g., `Lead_Shifted_Roll_{window}p_{shift}p`).
        *   *(Status: Done)*

    *   [x] **Testing:**
        *   Run the script with and without `--exclude-period`.
        *   Verify the scatter plot uses filtered data (if exclusions applied).
        *   Verify all other plots and the relevant Excel sheets ('Optimal Shift Data', 'Rolling Correlations', 'Cumulative Correlations') use the full, unfiltered data.
        *   Verify the bolding and bandwidth highlighting in the Excel sheets works correctly.
        *   Verify the 'R2 Results' sheet includes accurate final correlation values.
        *   *(Status: Pending)*

## 7. Known Issues & Bugs

*   Minor `FutureWarning` related to pandas concatenation with empty/NA entries during Excel export (appears benign for now).

---
*Self-Correction/Refinement during export implementation:* Added `max_shift` to `export_to_excel` to ensure correct formatting range for conditional highlighting, even if the optimal shift is outside the highlight band. Updated optimal data sheet to include shifts based on final rolling and cumulative correlations.