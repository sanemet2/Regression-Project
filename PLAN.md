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
*   More sophisticated missing data handling options.
*   Additional statistical tests (e.g., p-values for correlations).
*   Embedding plots directly into Excel (using `xlsxwriter` capabilities).
*   Automatically adjust time periods if one is weekly, the other monthly etc.
*   Fix charting functions and rolling correlation charts
*   Fix rolling correlations? Need to inquire as to why its mean reverting

## 6. Outlier Handling Plan (User-Specified Date Exclusion)

*   **Goal:** Allow the user to exclude specific date periods (e.g., COVID-19 shock) from the analysis to see the relationship based on more "normal" data.
*   **Implementation Steps:**
    *   [x] **Add CLI Argument:**
        *   Modify `main.py` `ArgumentParser`.
        *   Add a new argument `--exclude-period` using `action='append'`.
        *   Argument should accept a string format `YYYY-MM-DD:YYYY-MM-DD`.
        *   Include descriptive help text and `metavar`.
        *   *(Status: Argument added. Allows bypassing interactive prompt)*
    *   [x] **Add Interactive Prompting (Optional):**
        *   In `main.py`, before applying exclusions (Step 2.5), check if `args.exclude_period` is empty.
        *   If empty, ask the user "Do you want to specify any date periods to exclude? (y/n): ".
        *   If 'y':
            *   Implement a loop:
                *   Prompt for "Start date (YYYY-MM-DD) or leave blank to finish: ".
                *   If blank, break the loop.
                *   Prompt for "End date (YYYY-MM-DD): ".
                *   Validate date formats (YYYY-MM-DD).
                *   Validate start_date <= end_date.
                *   If valid, store the period string "START:END".
                *   If invalid, show an error and prompt again for that period.
            *   Store collected periods in a new list (e.g., `interactive_exclusions`).
        *   Modify the exclusion logic (Step 2.5) to process `interactive_exclusions` if `args.exclude_period` was empty but interactive periods were provided.
        *   Update summary print statements to show interactively added periods.
        *   *(Note: Command-line `--exclude-period` arguments will override interactive prompting).*
    *   [x] **Implement Filtering Logic:**
        *   In `main.py`, after `df = load_data(...)`.
        *   Check if `args.exclude_period` exists and is not empty.
        *   Initialize an empty boolean mask (e.g., `exclusion_mask = pd.Series(False, index=df.index)`).
        *   Loop through each `period_str` in `args.exclude_period`.
        *   Parse `start_date` and `end_date` from `period_str` (add error handling for incorrect format).
        *   Update the `exclusion_mask` to be `True` for dates within the current `start_date` and `end_date` range (`exclusion_mask = exclusion_mask | ((df.index >= start_date) & (df.index <= end_date))`).
        *   After the loop, filter the DataFrame: `df = df[~exclusion_mask]`.
        *   Print a message indicating how many rows were removed due to exclusions.
        *   *(Status: Logic implemented)*
    *   [x] **Documentation:**
        *   Briefly update the README (if one exists later) or add comments in `main.py` explaining the new arguments and functionality.
        *   *(Status: Section 2 & 3 of PLAN.md updated, comment added to main.py)*
    *   [x] **Testing:**
        *   Run the script *without* the `--exclude-period` argument and note the optimal shift/R-squared.
        *   Run the script *with* `--exclude-period` targeting the COVID-19 dates (e.g., `--exclude-period 2020-03-01:2020-09-01`) and verify:
            *   The exclusion message appears with the correct row counts.
            *   The resulting optimal shift and R-squared value change as expected (likely higher R-squared).
            *   The output plots and Excel file reflect the analysis run on the filtered data.
        *   Test with multiple `--exclude-period` arguments.
        *   Test with invalid date formats in `--exclude-period` to ensure error handling works (if implemented).
        *   *(Status: Testing complete, results verified)*
    *   [x] **Cleanup:** Remove the 'Winsorizing / trimming outliers' item from Section 5 ('Future Considerations') once this date exclusion feature is complete and tested.
        *   *(Status: Item removed from Section 5)*
