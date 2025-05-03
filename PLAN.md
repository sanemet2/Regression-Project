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
*   Fix charting functions and rolling correlation charts
*   Fix formatting

## 6. Excel Output Formatting Refinements

*   **Goal:** Ensure correct formatting (bold strongest column + highlight bandwidth) is applied to both 'Rolling Corrs' and 'Cumulative Corrs' sheets, and simplify the 'Optimal Shift Data' sheet.

*   **Debugging/Refinement Checklist:**
    1.  [x] **Review Formatting Function:** Check `_apply_correlation_formatting` in `export.py`. Verify parameters and logic for bolding (`apply_bolding`) and highlighting (`apply_highlighting`).
    2.  [x] **Inspect Strongest Shift Logic:** Check the code within `_apply_correlation_formatting` that determines `strongest_shift_col_name` and parses `strongest_shift_S`. Add print statements if necessary to verify values during execution.
    3.  [x] **Inspect Highlighting Logic (`apply_highlighting=True`):** (Investigated)
    4.  [x] **Inspect Bolding Logic (`apply_bolding=True`):** (Investigated)
    5.  [x] **Test Simplified Formatting:** (Investigated)
    6.  [x] **Implement Fixes:** Based on findings, modify the code in `export.py` to apply both bold and highlight to both correlation sheets.
    7.  [x] **Retest Formatting:** Run the script and verify the Excel output formatting is correct for both sheets.
    8.  [x] **Simplify 'Optimal Shift Data' Sheet:** Modify `export_to_excel` in `export.py` to remove the 'R-squared shift' column, keeping only target, best-rolling-shift, and best-cumulative-shift columns.
    9.  [x] **Retest 'Optimal Shift Data' Sheet:** Run script and verify the sheet simplification.

## 7. Known Issues & Bugs

*   Minor `FutureWarning` related to pandas concatenation with empty/NA entries during Excel export (appears benign for now).

---
*Self-Correction/Refinement during export implementation:* Added `max_shift` to `export_to_excel` to ensure correct formatting range for conditional highlighting, even if the optimal shift is outside the highlight band. Updated optimal data sheet to include shifts based on final rolling and cumulative correlations.