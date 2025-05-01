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
*   **Export Results:**
    *   Creates an output Excel file (`analysis_results.xlsx`).
    *   Includes sheets for:
        *   `R2 Results`: Table showing each tested lead/lag and its R-squared value.
        *   `Optimal Shift Data`: Columns for Date, Target Series, Leading Series (optimally shifted).
        *   `Rolling Correlations`: Columns for Date and the rolling correlation values for each tested lead/lag.
    *   Saves generated plots as separate image files.

## 3. Input Handling Strategy (CLI)

*   **File Path & Core Columns (`--file-path`, `--date-col`, `--leading-col`, `--target-col`):** These are optional command-line arguments. Default values (`C:\\Users\\franc\\OneDrive\\Desktop\\Programming\\Regression Project\\fredgraph.xlsx`, `date`, `icsa`, `unrate`, respectively) are hardcoded in `main.py` and shown in the help message (`-h`). The script uses these defaults if the arguments are not provided and does not prompt interactively for them.

*   **Analysis Parameters (`--range`, `--window`):** These are optional command-line arguments. If not provided via the command line, the script prompts the user interactively with descriptions of what the parameters mean. Basic positive integer validation is performed on the interactive input.

*   **Other Parameters (`--header`, `--sheet`, `--output_dir`):** These remain optional command-line arguments with defaults (`0`, `Monthly`, `results`, respectively). The script does not prompt interactively for these.

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
*   More sophisticated missing data handling options.
*   Additional statistical tests (e.g., p-values for correlations).
*   Embedding plots directly into Excel (using `xlsxwriter` capabilities).
*   Winsorizing / trimming outliers
*   Automatically adjust time periods if one is weekly, the other monthly etc.
*   Fix charting functions and rolling correlation charts

## 6. Code Refactoring Plan

*   **Efficiency & Conciseness in `analysis.py`:**
    *   [ ] Refactor `find_optimal_lead_lag` to potentially reduce complexity. Explore simplifying how the temporary DataFrame (`temp_df`) is handled inside the loop if possible without significant performance impact.
    *   [ ] Refactor `calculate_rolling_correlations` similarly. Look for ways to streamline the logic within the loop.
    *   [ ] **Test:** After refactoring `analysis.py`, run the script (`python main.py`, enter `24`, `24`) and compare output (optimal shift, R-squared, exported Excel) to ensure results are identical to the pre-refactoring run.
*   **Consolidation & Readability:**
    *   [ ] Review `plotting.py`: Identify opportunities to create helper functions for common plotting tasks (e.g., setting titles/labels, saving figures, standardizing appearance) to reduce code repetition and improve conciseness.
    *   [ ] Define Constants: Review all `.py` files for repeated strings (like internal column names 'Leading', 'Target', 'Shifted_Leading', 'Correlation_Shift_X') or magic numbers. Define these as constants (e.g., `LEADING_COL = 'Leading'`) at the top of the modules where they are primarily used (`analysis.py`, `plotting.py`, `export.py`).
    *   [ ] **Test:** After consolidating plotting code and defining constants, run the script again to ensure plots are generated correctly and results match previous runs.
*   **Minor Improvements & Cleanup:**
    *   [ ] Review interactive input loops in `main.py` for potential minor structural simplification.
    *   [ ] Enhance error handling in `data_loader.py` during date parsing (`pd.to_datetime`) to provide more specific feedback if formats are unexpected.
    *   [ ] (Low Priority) Address `FutureWarning` in `export.py` related to `pd.concat` by adjusting how the empty/NA row is handled, if feasible without complicating the logic significantly.
    *   [ ] **Test:** Perform a final end-to-end test run after all refactoring steps are complete using standard inputs.
