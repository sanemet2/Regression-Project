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

## 3. Technical Stack

*   **Language:** Python 3.x
*   **Core Libraries:**
    *   `pandas`: Data loading, manipulation, date handling, rolling calculations.
    *   `numpy`: Numerical operations.
    *   `scikit-learn`: Calculating R-squared (`sklearn.metrics.r2_score`).
    *   `matplotlib` / `seaborn`: Generating plots.
    *   `openpyxl` / `xlsxwriter`: Reading/writing Excel files.
    *   `argparse`: Handling command-line arguments.

## 4. Future Considerations

*   GUI Development (e.g., using Tkinter, PyQt, or Streamlit).
*   More sophisticated missing data handling options.
*   Additional statistical tests (e.g., p-values for correlations).
*   Embedding plots directly into Excel (using `xlsxwriter` capabilities).
*   Winsorizing / trimming outliers
*   Automatically adjust time periods if one is weekly, the other monthly etc.
*   Fix charting functions and rolling correlation charts

## 5. Input Handling Strategy (CLI)

*   **File Path & Core Columns (`--file-path`, `--date-col`, `--leading-col`, `--target-col`):**
    - [x] Implemented as optional command-line arguments.
    - [x] Default values are hardcoded in `main.py` (`Data/Input Data.xlsx`, `date`, `icsa`, `unrate`).
    - [x] Defaults are shown in the help message (`-h`).
    - [x] Script does *not* interactively prompt for these if missing; defaults are used.
*   **Analysis Parameters (`--range`, `--window`):**
    - [x] Implemented as optional command-line arguments.
    - [x] If not provided via command line, the script prompts the user interactively.
    - [x] Interactive prompts include descriptions of what the parameters mean.
    - [x] Basic validation (positive integer) is performed on interactive input.
*   **Other Parameters (`--header`, `--sheet`, `--output_dir`):**
    - [x] Remain optional command-line arguments with defaults.
    - [x] Script does not prompt interactively for these.
