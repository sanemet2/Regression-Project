# Project Plan: Time Series Lead/Lag Analysis App

## 1. Goal

Create a Python Command Line Interface (CLI) application that analyzes the lead/lag relationship between two time series from an Excel file. The app will identify the optimal lead/lag time based on R-squared, visualize the relationship, calculate rolling correlations for different leads/lags, and export results to Excel.

## 2. Core Features (Implemented in CLI)

*   **Data Loading:**
    *   Loads data from a user-specified Excel file (`.xlsx` or `.xls`).
    *   Allows the user to specify sheet name and header row.
    *   Handles single-frequency data using `--date-col`, `--leading-col`, `--target-col` (Though superseded by mixed-frequency handling below, kept for conceptual understanding).
    *   Parses dates, attempting common formats.
    *   Handles missing data by dropping rows with `NaN` in selected columns.
    *   Optionally excludes user-specified date periods (e.g., `--exclude-period START:END`) before analysis.
*   **Mixed-Frequency Handling:**
    *   Supports input data with separate date and value columns for leading and target series (using `--lead-date-col`, `--lead-val-col`, `--target-date-col`, `--target-val-col`).
    *   Automatically infers the frequency (e.g., Daily, Weekly, Monthly, Quarterly) for each series based on its date column.
    *   If frequencies differ, automatically downsamples the higher-frequency series to match the lower-frequency one.
        *   Currently uses the *last* observation within the lower frequency period for resampling (e.g., last weekly value represents the month).
        *   Handles specific Weekly vs. Monthly resampling logic.
    *   Performs all subsequent analysis on the aligned, potentially resampled data.
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

*   **File Path (`--file-path`):** Optional command-line argument. Default: `C:\Users\franc\OneDrive\Desktop\Programming\Regression Project\fredgraph.xlsx`. Not prompted interactively.
*   **Core Columns (Mixed Frequency):**
    *   `--lead-date-col` (Default: `observation_date`)
    *   `--lead-val-col` (Default: `ICSA`)
    *   `--target-date-col` (Default: `date`)
    *   `--target-val-col` (Default: `unrate`)
    *   These are optional arguments with defaults shown in help (`-h`). They replace the older single `--date-col`, `--leading-col`, `--target-col` structure. Not prompted interactively.
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
*   Changing series to year over year, BP change, etc. 
*   How will sampled data be represented in the chart? Original? Resamopled? Last point averaged?
*   How do we detect input data? 



## 6. GUI Development (Streamlit)

**Goal:** Provide a user-friendly interface for loading Excel workbooks, selecting date/value columns, configuring analysis parameters, running the mixed-frequency analysis, viewing results interactively, exporting reports, and logging inputs.

### Implementation Checklist:

1. [x] Add new dependency `streamlit>=1.0` to `requirements.txt` and install it.
2. [x] Create `app.py` at the project root. Import `streamlit as st` and core modules (`load_data`, `analysis`, `export_results`).
3. [x] Implement file uploader widget to load `.xls`/`.xlsx` files and display available sheet names.
   - [x] Test: uploader appears and sheet names list correctly.
4. [x] Read selected sheet with `pandas.read_excel`, show a preview table via `st.dataframe`, and drop rows with NaNs.
   - [x] Test: DataFrame preview displays correctly.
5. [x] Dynamically populate dropdowns for the four columns: lead date, lead value, target date, target value. Ensure date parsing handles `mm/yy` format.
   - [x] Test: Column selection dropdowns appear and are populated.
6. [x] Add widgets for parameters: `max_shift` (int), `window` (int), `exclude_period` (text `START:END`).
   - [x] Test: Parameter input widgets appear correctly.
7. [x] On “Run Analysis” button click:
   - [x] Call `load_data(...)` with the uploaded file and selected columns.
   - [x] Drop missing data and apply exclude periods.
   - [x] Test: raw series loaded without errors.
8. [x] Detect frequencies with `pd.infer_freq`; if differ, resample the higher-frequency series (`.resample(...).last()`), then align series into a DataFrame.
9. [x] Call analysis functions: `find_optimal_lead_lag`, `calculate_rolling_correlations`, `calculate_cumulative_correlations`.

> ✅ We will check off each box as we complete the step.

## 7. Known Issues & Bugs

*   Minor `FutureWarning` related to pandas concatenation with empty/NA entries during Excel export (appears benign for now).

---
*Self-Correction/Refinement during export implementation:* Added `max_shift` to `export_to_excel` to ensure correct formatting range for conditional highlighting, even if the optimal shift is outside the highlight band. Updated optimal data sheet to include shifts based on final rolling and cumulative correlations.